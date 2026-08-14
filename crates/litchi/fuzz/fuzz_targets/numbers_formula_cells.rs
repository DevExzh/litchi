#![no_main]

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    cell::Value,
    formula::{
        AxisReference, BinaryOperator, CachedValue, CellReference, Error as FormulaError,
        Expression,
    },
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
const MAX_FORMULA_WORK: usize = 8 * 1024;
const MAX_FORMULA_DEPTH: usize = 32;
const MAX_FORMULA_INPUT_BYTES: usize = 256;
const CONTROL_BYTES: usize = 8;
const PRIVATE_SELECTOR: &str = "__litchi_private_formula_selector_65d8__";
const PRIVATE_TEXT: &str = "__litchi_private_formula_text_65d8__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

fuzz_target!(|data: &[u8]| {
    // Keep malformed ZIP/IWA ingress under the same finite profile as the
    // other Numbers targets. Most arbitrary bytes fail CRC admission, so the
    // native seed below also receives every input as a semantic command.
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_reads(&package, data),
        Err(error) => observe_error(error),
    }

    exercise_reads(native_package(), data);
    exercise_formula_constructors(data);
    exercise_foreign_handle(data);
    exercise_private_errors(native_package());
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
            .unwrap_or_else(|error| panic!("native Numbers formula seed must open: {error}"));
        let table = package
            .document()
            .sheets()
            .first()
            .and_then(|sheet| sheet.tables().next())
            .unwrap_or_else(|| panic!("native Numbers formula seed must expose a table"));
        assert!(table.dimensions().rows() > 0);
        assert!(table.dimensions().columns() > 0);
        package
    })
}

fn foreign_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_NUMBERS, fuzz_options())
            .unwrap_or_else(|error| panic!("foreign Numbers formula seed must open: {error}"))
    })
}

fn exercise_reads(package: &Package, data: &[u8]) {
    let Some(table) = package
        .document()
        .sheets()
        .first()
        .and_then(|sheet| sheet.tables().next())
    else {
        return;
    };
    let dimensions = table.dimensions();
    let position = position(data, 2, dimensions);
    observe_result(package.table_cell(0usize, 0usize, position));
    observe_result(package.table_cells(
        0usize,
        0usize,
        CellRange::single(position).unwrap_or_else(|error| {
            panic!("a checked native cell position must form a range: {error}")
        }),
    ));
    exercise_batch(package, data, dimensions);
}

fn exercise_batch(package: &Package, data: &[u8], dimensions: Dimensions) {
    let target = position(data, 0, dimensions);
    let formula_target = position(data, 1, dimensions);
    let command = control(data, 0) % 8;
    let before = exact_bytes(package);

    let result = match command {
        0 => no_op_batch(package),
        1 => changed_formula_batch(package, data, dimensions, target, formula_target),
        2 => cached_formula_batch(package, data, dimensions, formula_target),
        3 => mismatched_cache_batch(package, data, dimensions, formula_target),
        4 => clear_batch(package, formula_target),
        5 => duplicate_batch(package, formula_target),
        6 => cycle_batch(package, dimensions, formula_target),
        _ => mixed_batch(package, data, dimensions, target, formula_target),
    };

    match result {
        Ok(commit) => assert_commit_round_trip(package, &before, commit, formula_target),
        Err(error) => {
            observe_error(&error);
            assert_eq!(exact_bytes(package), before);
        },
    }
}

fn no_op_batch(package: &Package) -> Result<litchi::numbers::table::cells::Commit, CellError> {
    let position = CellPosition::new(2, 1);
    let Ok(state) = package.table_cell(0usize, 0usize, position) else {
        return package
            .edit_table_cells(0usize, 0usize)?
            .clear(position)
            .and_then(|edit| edit.commit());
    };
    let Storage::Stored(Value::Number(value)) = state.storage() else {
        return package
            .edit_table_cells(0usize, 0usize)?
            .clear(position)
            .and_then(|edit| edit.commit());
    };
    let input = Input::number(value.get()).map_err(|_| CellError::InvalidSource {
        path: litchi::numbers::table::cells::Path::Package,
    })?;
    package
        .edit_table_cells(0usize, 0usize)?
        .set(position, input)?
        .commit()
}

fn changed_formula_batch(
    package: &Package,
    data: &[u8],
    dimensions: Dimensions,
    target: CellPosition,
    formula_target: CellPosition,
) -> Result<litchi::numbers::table::cells::Commit, CellError> {
    let expression = expression_for(package, data, dimensions, formula_target)?;
    let scalar =
        Input::number(f64::from(control(data, 3)) + 1.0).map_err(|_| CellError::InvalidSource {
            path: litchi::numbers::table::cells::Path::Package,
        })?;
    package
        .edit_table_cells(0usize, 0usize)?
        .set(target, scalar)?
        .set_formula(formula_target, expression)?
        .commit()
}

fn cached_formula_batch(
    package: &Package,
    data: &[u8],
    dimensions: Dimensions,
    formula_target: CellPosition,
) -> Result<litchi::numbers::table::cells::Commit, CellError> {
    let expression = simple_sum_expression(data, dimensions);
    let cache = CachedValue::number(42.0).map_err(|_| CellError::InvalidSource {
        path: litchi::numbers::table::cells::Path::Package,
    })?;
    package
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached(formula_target, expression, cache)?
        .commit()
}

fn mismatched_cache_batch(
    package: &Package,
    data: &[u8],
    dimensions: Dimensions,
    formula_target: CellPosition,
) -> Result<litchi::numbers::table::cells::Commit, CellError> {
    let expression = simple_sum_expression(data, dimensions);
    let cache = CachedValue::number(999.0).map_err(|_| CellError::InvalidSource {
        path: litchi::numbers::table::cells::Path::Package,
    })?;
    package
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached(formula_target, expression, cache)?
        .commit()
}

fn clear_batch(
    package: &Package,
    formula_target: CellPosition,
) -> Result<litchi::numbers::table::cells::Commit, CellError> {
    package
        .edit_table_cells(0usize, 0usize)?
        .clear(formula_target)?
        .commit()
}

fn duplicate_batch(
    package: &Package,
    formula_target: CellPosition,
) -> Result<litchi::numbers::table::cells::Commit, CellError> {
    let edit = package
        .edit_table_cells(0usize, 0usize)?
        .clear(formula_target)?
        .clear(formula_target)?;
    edit.commit()
}

fn cycle_batch(
    package: &Package,
    dimensions: Dimensions,
    formula_target: CellPosition,
) -> Result<litchi::numbers::table::cells::Commit, CellError> {
    let other = CellPosition::new(
        formula_target.row(),
        (formula_target.column() + 1) % dimensions.columns().max(1),
    );
    let first = Expression::cell(CellReference::relative(
        other.row() as usize,
        other.column() as usize,
    ));
    let second = Expression::cell(CellReference::relative(
        formula_target.row() as usize,
        formula_target.column() as usize,
    ));
    package
        .edit_table_cells(0usize, 0usize)?
        .set_formula(formula_target, first)?
        .set_formula(other, second)?
        .commit()
}

fn mixed_batch(
    package: &Package,
    data: &[u8],
    dimensions: Dimensions,
    target: CellPosition,
    formula_target: CellPosition,
) -> Result<litchi::numbers::table::cells::Commit, CellError> {
    let expression = expression_for(package, data, dimensions, formula_target)?;
    let text = input_text(data);
    package
        .edit_table_cells(0usize, 0usize)?
        .set(target, Input::text(text)?)?
        .set_formula(formula_target, expression)?
        .commit()
}

fn assert_commit_round_trip(
    package: &Package,
    source: &[u8],
    commit: litchi::numbers::table::cells::Commit,
    formula_target: CellPosition,
) {
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    let target = exact_bytes(commit.package());
    assert_eq!(patch.is_noop(), target == source);
    assert_eq!(diagnostics.changed(), !patch.is_noop());
    assert_eq!(diagnostics.full_reparse_performed(), !patch.is_noop());
    assert_eq!(diagnostics.requested_cells(), patch.len());
    if patch.is_noop() {
        assert_eq!(diagnostics.touched_components(), 0);
        assert_eq!(diagnostics.refreshed_formula_caches(), 0);
    } else {
        assert!(diagnostics.touched_components() > 0);
    }
    if let Ok(state) = commit.package().table_cell(0usize, 0usize, formula_target) {
        black_box(state);
    }

    let applied = package
        .apply_table_cells(&patch)
        .unwrap_or_else(|error| panic!("fresh formula-cell patch must apply: {error}"));
    assert_eq!(exact_bytes(applied.package()), target);
    if !patch.is_noop() {
        assert!(matches!(
            applied.package().apply_table_cells(&patch),
            Err(CellError::PatchConflict)
        ));
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_table_cells(&inverse)
        .unwrap_or_else(|error| panic!("fresh formula-cell inverse must apply: {error}"));
    assert_eq!(exact_bytes(restored.package()), source);
}

fn expression_for(
    package: &Package,
    data: &[u8],
    dimensions: Dimensions,
    target: CellPosition,
) -> Result<Expression, CellError> {
    let left = simple_sum_expression(data, dimensions);
    let right = Expression::number(f64::from(control(data, 5) % 9) + 1.0).map_err(|_| {
        CellError::InvalidSource {
            path: litchi::numbers::table::cells::Path::Package,
        }
    })?;
    let expression = match control(data, 4) % 7 {
        0 => Ok(left),
        1 => Expression::binary(BinaryOperator::Add, left, right),
        2 => Expression::binary(BinaryOperator::Multiply, left, right),
        3 => Expression::negate(left),
        4 => Expression::percent(left),
        5 => {
            let table = package
                .edit_table_cells(0usize, 0usize)?
                .formula_table(0usize, 0usize)?;
            Ok(Expression::table_cell(
                &table,
                CellReference::relative(target.row() as usize, target.column() as usize),
            ))
        },
        _ => {
            let argument = Expression::cell(CellReference::relative(
                target.row() as usize,
                target.column() as usize,
            ));
            Expression::function("ABS", [argument])
        },
    };
    expression.map_err(|_| CellError::InvalidSource {
        path: litchi::numbers::table::cells::Path::Package,
    })
}

fn simple_sum_expression(data: &[u8], dimensions: Dimensions) -> Expression {
    let row = usize::from(control(data, 1)) % dimensions.rows().max(1) as usize;
    let column = usize::from(control(data, 2)) % dimensions.columns().max(1) as usize;
    Expression::function(
        "SUM",
        [
            Expression::cell(CellReference::relative(row, column)),
            Expression::number(f64::from(control(data, 6) % 11))
                .unwrap_or_else(|error| unreachable!("finite fuzz literal: {error}")),
        ],
    )
    .unwrap_or_else(|error| unreachable!("valid SUM expression: {error}"))
}

fn exercise_formula_constructors(data: &[u8]) {
    for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        assert!(matches!(
            Expression::number(value),
            Err(FormulaError::NonFinite)
        ));
        assert!(matches!(
            CachedValue::number(value),
            Err(FormulaError::NonFinite)
        ));
    }
    observe_result(Expression::text(input_text(data)));
    observe_result(Expression::function(
        "NOT",
        [Expression::boolean(control(data, 7) & 1 != 0)],
    ));
    observe_result(Expression::range(
        CellReference::relative(0, 0),
        CellReference::absolute(1, 1),
    ));
    observe_result(Expression::rows(
        AxisReference::relative(0),
        AxisReference::absolute(1),
    ));
    observe_result(Expression::columns(
        AxisReference::relative(0),
        AxisReference::absolute(1),
    ));
}

fn exercise_foreign_handle(data: &[u8]) {
    let source = native_package();
    let foreign_table = foreign_package()
        .edit_table_cells(0usize, 0usize)
        .unwrap_or_else(|error| panic!("foreign formula edit must resolve: {error}"))
        .formula_table(0usize, 0usize)
        .unwrap_or_else(|error| panic!("foreign formula table must resolve: {error}"));
    let expression = Expression::table_cell(
        &foreign_table,
        CellReference::relative(usize::from(control(data, 1)) % 4, 0),
    );
    let result = source
        .edit_table_cells(0usize, 0usize)
        .and_then(|edit| edit.set_formula(CellPosition::new(2, 2), expression));
    assert!(matches!(result, Err(CellError::PatchConflict)));
}

fn exercise_private_errors(package: &Package) {
    if let Err(error) = package.table_cell(
        SheetSelector::name(PRIVATE_SELECTOR),
        TableSelector::index(0),
        CellPosition::new(0, 0),
    ) {
        observe_redacted(error, PRIVATE_SELECTOR);
    }
    if let Err(error) = package
        .edit_table_cells(0usize, 0usize)
        .and_then(|edit| edit.set_formula_a1("private formula address", Expression::boolean(true)))
    {
        observe_error(error);
    }
    black_box(PRIVATE_TEXT);
}

fn position(data: &[u8], offset: usize, dimensions: Dimensions) -> CellPosition {
    let rows = dimensions.rows().max(1);
    let columns = dimensions.columns().max(1);
    CellPosition::new(
        u32::from(control(data, offset)) % rows,
        u32::from(control(data, offset + 1)) % columns,
    )
}

fn input_text(data: &[u8]) -> String {
    let start = data.len().min(CONTROL_BYTES);
    let end = data
        .len()
        .min(start.saturating_add(MAX_FORMULA_INPUT_BYTES));
    String::from_utf8_lossy(&data[start..end]).into_owned()
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn exact_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Numbers package to memory must succeed: {error}"));
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
