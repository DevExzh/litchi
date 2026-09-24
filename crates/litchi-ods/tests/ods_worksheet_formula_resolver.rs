//! Integration coverage for the borrowed worksheet formula resolver.
//!
//! These tests exercise the adapter at its public boundary.  The worksheet
//! model retains ODF repetition as physical runs, so the fixture deliberately
//! uses repeated rows and cells and checks logical-coordinate lookup without
//! expanding those runs or copying their values.

use std::{
    num::{NonZeroU64, NonZeroUsize},
    ptr,
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits,
    Limits as CoreLimits, Profile, Resource,
};
use litchi_ods::{
    Builder, Cell, CellValue, Row, Sheet,
    codec::formula::evaluation::{
        EvaluationFailure, ScalarError, UnsupportedKind,
        value::{
            CellRead, Context, Evaluated, Limits as ValueLimits, Mode, Position,
            Resolver as ValueResolver, SheetExtent, Value, evaluate,
        },
    },
    codec::formula::expression::Expression,
    worksheet::{
        Snapshot,
        formula::{Error as FormulaError, Resolver},
    },
};

fn make_execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    execution_with_limits(scope, CoreLimits::for_profile(Profile::Server))
}

fn execution_with_limits(
    scope: &str,
    limits: CoreLimits,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), limits);
    let (cancellation, token) = CancellationSource::pair();
    let execution_limits =
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)
            .expect("one-worker execution policy is valid");
    (
        budget.clone(),
        cancellation,
        ExecutionContext::new(budget, token, execution_limits),
    )
}

fn repeated_sheet() -> Sheet {
    let mut sheet = Sheet::new("Data").expect("sheet name is valid");

    let mut repeated = Row::repeated(2).expect("positive row repetition");
    repeated
        .push_cell(
            Cell::repeated(CellValue::Text("repeated text".to_owned()), "display", 3)
                .expect("positive cell repetition"),
        )
        .expect("valid repeated text cell");
    repeated
        .push_cell(Cell::new(CellValue::Number(7.5), "7.5"))
        .expect("valid number cell");
    sheet.push_row(repeated).expect("valid repeated row");

    let mut tail = Row::new();
    tail.push_cell(Cell::new(CellValue::Text("tail".to_owned()), "tail"))
        .expect("valid tail cell");
    sheet.push_row(tail).expect("valid tail row");
    sheet
}

fn read<'resolver>(
    resolver: &'resolver Resolver<'_>,
    execution: &ExecutionContext,
    sheet: &str,
    row: usize,
    column: usize,
) -> CellRead<'resolver> {
    ValueResolver::read_cell(resolver, sheet, row, column, execution)
        .expect("worksheet reads do not fail operationally")
}

fn evaluate_worksheet<'expr>(
    expression: &'expr Expression,
    resolver: &'expr Resolver<'_>,
    execution: &ExecutionContext,
    sheet: &str,
    row: usize,
    column: usize,
    mode: Mode,
) -> Result<Evaluated<'expr>, EvaluationFailure> {
    let context = Context::new(execution, Position::new(sheet, row, column)).with_mode(mode);
    evaluate(expression, resolver, &context, &ValueLimits::default())
}

#[test]
fn logical_repeated_rows_and_cells_resolve_at_run_boundaries() {
    let sheets = vec![
        repeated_sheet(),
        Sheet::new("Other").expect("sheet name is valid"),
    ];
    let (_budget, _cancellation, execution) = make_execution("worksheet-formula-runs");
    let resolver = Resolver::new(&sheets, SheetExtent::new(4, 5), &execution)
        .expect("repeated worksheet index should build");

    assert_eq!(
        ValueResolver::sheet_extent(&resolver, "Data", &execution).unwrap(),
        Some(SheetExtent::new(4, 5))
    );
    assert_eq!(
        ValueResolver::sheet_extent(&resolver, "Missing", &execution).unwrap(),
        None
    );
    assert_eq!(
        ValueResolver::sheet_index(&resolver, "Data", &execution).unwrap(),
        Some(0)
    );
    assert_eq!(
        ValueResolver::sheet_index(&resolver, "Missing", &execution).unwrap(),
        None
    );
    assert_eq!(
        ValueResolver::sheet_name_at(&resolver, 0, &execution).unwrap(),
        Some("Data")
    );
    assert_eq!(
        ValueResolver::sheet_name_at(&resolver, 1, &execution).unwrap(),
        Some("Other")
    );
    assert_eq!(
        ValueResolver::sheet_name_at(&resolver, 2, &execution).unwrap(),
        None
    );
    assert_eq!(
        ValueResolver::sheet_count(&resolver, &execution).unwrap(),
        2
    );

    for row in [0, 1] {
        for column in [0, 1, 2] {
            assert!(matches!(
                read(&resolver, &execution, "Data", row, column),
                CellRead::Text(text) if text == "repeated text"
            ));
        }
        assert!(matches!(
            read(&resolver, &execution, "Data", row, 3),
            CellRead::Number(value) if value == 7.5
        ));
    }

    assert!(matches!(
        read(&resolver, &execution, "Data", 2, 0),
        CellRead::Text(text) if text == "tail"
    ));
    // A physical row with no cells and a coordinate past the stored run are
    // both ordinary empty cells while they remain inside the caller extent.
    assert!(matches!(
        read(&resolver, &execution, "Data", 3, 0),
        CellRead::Empty
    ));
    assert!(matches!(
        read(&resolver, &execution, "Data", 0, 4),
        CellRead::Empty
    ));
}

#[test]
fn text_reads_borrow_the_cell_value_without_copying() {
    let sheets = vec![repeated_sheet()];
    let source_text = match &sheets[0].rows[0].cells[0].value {
        CellValue::Text(text) => text.as_str(),
        value => panic!("fixture text cell unexpectedly contains {value:?}"),
    };
    let source_ptr = source_text.as_ptr();
    let (_budget, _cancellation, execution) = make_execution("worksheet-formula-text-borrow");
    let resolver = Resolver::new(&sheets, SheetExtent::new(4, 5), &execution)
        .expect("worksheet index should build");

    let first = read(&resolver, &execution, "Data", 0, 0);
    let second = read(&resolver, &execution, "Data", 1, 2);
    let (CellRead::Text(first), CellRead::Text(second)) = (first, second) else {
        panic!("repeated text cells should produce borrowed text reads")
    };
    assert_eq!(first, source_text);
    assert_eq!(second, source_text);
    assert!(ptr::eq(first.as_ptr(), source_ptr));
    assert!(ptr::eq(second.as_ptr(), source_ptr));
}

#[test]
fn missing_and_out_of_grid_coordinates_are_formula_reference_errors() {
    let sheets = vec![repeated_sheet()];
    let (_budget, _cancellation, execution) = make_execution("worksheet-formula-reference-errors");
    let resolver = Resolver::new(&sheets, SheetExtent::new(4, 5), &execution)
        .expect("worksheet index should build");

    for (sheet, row, column) in [
        ("Missing", 0, 0),
        ("Data", 4, 0),
        ("Data", 0, 5),
        ("Data", usize::MAX, 0),
        ("Data", 0, usize::MAX),
    ] {
        assert!(
            matches!(
                read(&resolver, &execution, sheet, row, column),
                CellRead::Error(ScalarError::Reference)
            ),
            "{sheet}[{row},{column}] should be #REF!"
        );
    }
}

#[test]
fn supported_cell_types_map_to_value_resolver_reads() {
    let mut sheet = Sheet::new("Types").expect("sheet name is valid");
    let mut row = Row::new();
    for cell in [
        Cell::new(CellValue::Empty, ""),
        Cell::new(CellValue::Number(2.25), "2.25"),
        Cell::new(
            CellValue::Currency {
                value: 3.5,
                currency: "USD".to_owned(),
            },
            "$3.50",
        ),
        Cell::new(CellValue::Percentage(0.125), "12.5%"),
        Cell::new(CellValue::Boolean(true), "TRUE"),
    ] {
        row.push_cell(cell).expect("supported cell is valid");
    }
    sheet.push_row(row).expect("supported row is valid");
    let sheets = vec![sheet];
    let (_budget, _cancellation, execution) = make_execution("worksheet-formula-types");
    let resolver = Resolver::new(&sheets, SheetExtent::new(1, 5), &execution)
        .expect("worksheet index should build");

    assert!(matches!(
        read(&resolver, &execution, "Types", 0, 0),
        CellRead::Empty
    ));
    assert!(matches!(
        read(&resolver, &execution, "Types", 0, 1),
        CellRead::Number(value) if value == 2.25
    ));
    assert!(matches!(
        read(&resolver, &execution, "Types", 0, 2),
        CellRead::Number(value) if value == 3.5
    ));
    assert!(matches!(
        read(&resolver, &execution, "Types", 0, 3),
        CellRead::Number(value) if value == 0.125
    ));
    assert!(matches!(
        read(&resolver, &execution, "Types", 0, 4),
        CellRead::Logical(true)
    ));
}

#[test]
fn formula_cells_refuse_cached_values_before_value_dispatch() {
    let mut sheet = Sheet::new("FormulaCache").expect("sheet name is valid");
    let mut row = Row::new();
    let mut number = Cell::new(CellValue::Number(11.0), "11");
    number.set_formula("of:=11").expect("formula is valid");
    let mut text = Cell::new(CellValue::Text("cached".to_owned()), "cached");
    text.set_formula("of:=\"cached\"")
        .expect("formula is valid");
    let mut logical = Cell::new(CellValue::Boolean(true), "TRUE");
    logical.set_formula("of:=TRUE()").expect("formula is valid");
    let mut empty = Cell::empty();
    empty.set_formula("of:=0").expect("formula is valid");
    let mut currency = Cell::new(
        CellValue::Currency {
            value: 12.0,
            currency: "USD".to_owned(),
        },
        "$12",
    );
    currency.set_formula("of:=12").expect("formula is valid");
    let mut percentage = Cell::new(CellValue::Percentage(0.25), "25%");
    percentage.set_formula("of:=25%").expect("formula is valid");
    let mut date = Cell::new(CellValue::Date("2026-09-13".to_owned()), "date");
    date.set_formula("of:=0").expect("formula is valid");
    let mut time = Cell::new(CellValue::Time("PT1H".to_owned()), "time");
    time.set_formula("of:=0").expect("formula is valid");
    let mut unknown = Cell::new(
        CellValue::Unknown {
            kind: "future-type".to_owned(),
            value: Some("opaque".to_owned()),
        },
        "opaque",
    );
    unknown.set_formula("of:=0").expect("formula is valid");
    for cell in [
        number, text, logical, empty, currency, percentage, date, time, unknown,
    ] {
        row.push_cell(cell).expect("formula cache cell is valid");
    }
    sheet.push_row(row).expect("formula cache row is valid");
    let sheets = vec![sheet];
    let (_budget, _cancellation, execution) = make_execution("worksheet-formula-cache");
    let resolver = Resolver::new(&sheets, SheetExtent::new(1, 9), &execution)
        .expect("worksheet index should build");

    for column in 0..9 {
        assert!(
            matches!(
                read(&resolver, &execution, "FormulaCache", 0, column),
                CellRead::Unsupported
            ),
            "formula cache at column {column} must not be treated as authoritative"
        );
    }
}

#[test]
fn unsupported_date_time_and_unknown_values_are_refused() {
    let mut sheet = Sheet::new("Unsupported").expect("sheet name is valid");
    let mut row = Row::new();
    for cell in [
        Cell::new(CellValue::Date("2026-09-13".to_owned()), "date"),
        Cell::new(CellValue::Time("PT1H".to_owned()), "time"),
        Cell::new(
            CellValue::Unknown {
                kind: "future-type".to_owned(),
                value: Some("opaque".to_owned()),
            },
            "opaque",
        ),
    ] {
        row.push_cell(cell)
            .expect("unsupported stored value is valid");
    }
    sheet.push_row(row).expect("unsupported row is valid");
    let sheets = vec![sheet];
    let (_budget, _cancellation, execution) = make_execution("worksheet-formula-unsupported");
    let resolver = Resolver::new(&sheets, SheetExtent::new(1, 3), &execution)
        .expect("worksheet index should build");

    for column in 0..3 {
        assert!(matches!(
            read(&resolver, &execution, "Unsupported", 0, column),
            CellRead::Unsupported
        ));
    }
}

#[test]
fn nonfinite_numbers_become_formula_number_errors() {
    // Push directly into the public model fields to exercise the resolver's
    // defensive dispatch.  Normal authoring APIs reject non-finite numbers,
    // but an adapter must still avoid exposing NaN or infinity if it receives
    // a malformed in-memory graph from another producer.
    let mut sheet = Sheet::new("NonFinite").expect("sheet name is valid");
    let mut row = Row::new();
    row.cells
        .push(Cell::new(CellValue::Number(f64::NAN), "NaN"));
    row.cells
        .push(Cell::new(CellValue::Number(f64::INFINITY), "infinity"));
    sheet.rows.push(row);
    let sheets = vec![sheet];
    let (_budget, _cancellation, execution) = make_execution("worksheet-formula-nonfinite");
    let resolver = Resolver::new(&sheets, SheetExtent::new(1, 2), &execution)
        .expect("indexing does not silently convert non-finite values");

    for column in 0..2 {
        assert!(matches!(
            read(&resolver, &execution, "NonFinite", 0, column),
            CellRead::Error(ScalarError::Number)
        ));
    }
}

#[test]
fn duplicate_sheet_names_fail_atomically_and_release_index_memory() {
    let sheets = vec![
        Sheet::new("Duplicate").expect("sheet name is valid"),
        Sheet::new("Duplicate").expect("sheet name is valid"),
    ];
    let (budget, _cancellation, execution) = make_execution("worksheet-formula-duplicate");
    let result = Resolver::new(&sheets, SheetExtent::new(1, 1), &execution);
    assert!(matches!(result, Err(FormulaError::DuplicateSheetName)));
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn invalid_extent_and_cancellation_leave_budget_unchanged() {
    let sheets = vec![repeated_sheet()];

    let (budget, _cancellation, execution) = make_execution("worksheet-formula-invalid-extent");
    assert!(matches!(
        Resolver::new(&sheets, SheetExtent::new(0, 4), &execution),
        Err(FormulaError::InvalidExtent)
    ));
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation, execution) = make_execution("worksheet-formula-cancelled");
    cancellation.cancel();
    assert!(matches!(
        Resolver::new(&sheets, SheetExtent::new(4, 4), &execution),
        Err(FormulaError::Execution(ExecutionError::Cancelled))
    ));
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn memory_and_work_admission_fail_without_partial_index_reservations() {
    let sheets = vec![repeated_sheet()];

    let (budget, _cancellation, execution) = execution_with_limits(
        "worksheet-formula-memory-limit",
        CoreLimits::new(0, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    assert!(matches!(
        Resolver::new(&sheets, SheetExtent::new(4, 4), &execution),
        Err(FormulaError::ResourceLimit(limit)) if limit.resource == Resource::Memory
    ));
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, execution) = execution_with_limits(
        "worksheet-formula-work-limit",
        CoreLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, 1),
    );
    assert!(matches!(
        Resolver::new(&sheets, SheetExtent::new(4, 4), &execution),
        Err(FormulaError::ResourceLimit(limit)) if limit.resource == Resource::Work
    ));
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn post_construction_metadata_lookups_honor_cancellation_and_release_index_memory() {
    let sheets = vec![
        repeated_sheet(),
        Sheet::new("Other").expect("sheet name is valid"),
    ];
    let (budget, cancellation, execution) = make_execution("worksheet-formula-post-cancel");
    let resolver = Resolver::new(&sheets, SheetExtent::new(4, 5), &execution)
        .expect("worksheet index should build before cancellation");
    assert!(resolver.reserved_index_bytes() > 0);
    assert_eq!(
        budget.used(Resource::Memory),
        resolver.reserved_index_bytes()
    );

    cancellation.cancel();
    assert!(matches!(
        ValueResolver::sheet_count(&resolver, &execution),
        Err(EvaluationFailure::Cancelled)
    ));
    assert!(matches!(
        ValueResolver::sheet_name_at(&resolver, 0, &execution),
        Err(EvaluationFailure::Cancelled)
    ));
    assert!(matches!(
        ValueResolver::read_cell(&resolver, "Data", 0, 0, &execution),
        Err(EvaluationFailure::Cancelled)
    ));

    drop(resolver);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn lookups_use_the_caller_execution_policy_after_independent_preparation() {
    let sheets = vec![repeated_sheet()];
    let (preparation_budget, _preparation_cancellation, preparation_execution) =
        make_execution("worksheet-formula-independent-preparation");
    let resolver = Resolver::new(&sheets, SheetExtent::new(4, 5), &preparation_execution)
        .expect("independent preparation should build the worksheet index");
    assert_eq!(
        preparation_budget.used(Resource::Memory),
        resolver.reserved_index_bytes()
    );

    // A later caller with no lookup work must be rejected by the lookup's
    // supplied context, even though preparation had a separate budget.
    let (_caller_budget, _caller_cancellation, constrained_execution) = execution_with_limits(
        "worksheet-formula-independent-constrained-caller",
        CoreLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, 0),
    );
    assert!(matches!(
        ValueResolver::sheet_count(&resolver, &constrained_execution),
        Err(EvaluationFailure::ResourceLimit(limit)) if limit.resource == Resource::Work
    ));

    // Evaluation forwards its caller context to the resolver as well.  A
    // cancelled caller must stop before reading a cell from the prepared
    // index.
    let (_caller_budget, caller_cancellation, caller_execution) =
        make_execution("worksheet-formula-independent-cancelled-caller");
    caller_cancellation.cancel();
    let expression = Expression::parse("=[.A1]").expect("reference should parse");
    let error = evaluate_worksheet(
        &expression,
        &resolver,
        &caller_execution,
        "Data",
        0,
        0,
        Mode::Scalar,
    )
    .expect_err("cancelled caller policy must reach the worksheet lookup");
    assert!(matches!(error, EvaluationFailure::Cancelled));

    drop(resolver);
    assert_eq!(preparation_budget.used(Resource::Memory), 0);
}

#[test]
fn preparation_cancellation_does_not_cancel_a_healthy_independent_lookup() {
    let sheets = vec![repeated_sheet()];
    let (preparation_budget, preparation_cancellation, preparation_execution) =
        make_execution("worksheet-formula-preparation-cancellation");
    let resolver = Resolver::new(&sheets, SheetExtent::new(4, 5), &preparation_execution)
        .expect("preparation should complete before its token is cancelled");
    let reserved = resolver.reserved_index_bytes();
    assert_eq!(preparation_budget.used(Resource::Memory), reserved);

    preparation_cancellation.cancel();
    let (_caller_budget, _caller_cancellation, caller_execution) =
        make_execution("worksheet-formula-healthy-independent-caller");
    assert!(matches!(
        read(&resolver, &caller_execution, "Data", 0, 3),
        CellRead::Number(value) if value == 7.5
    ));
    assert_eq!(
        ValueResolver::sheet_count(&resolver, &caller_execution)
            .expect("healthy caller should use the prepared index"),
        1
    );
    assert_eq!(
        preparation_budget.used(Resource::Memory),
        reserved,
        "index reservation belongs to preparation until resolver drop"
    );

    drop(resolver);
    assert_eq!(preparation_budget.used(Resource::Memory), 0);
}

#[test]
fn from_snapshot_reuses_the_parsed_worksheet_graph() {
    let sheet = repeated_sheet();
    let mut builder = Builder::new();
    builder
        .add_sheet(sheet)
        .expect("worksheet fixture should be authorable");
    let snapshot =
        Snapshot::from_bytes(builder.build().expect("worksheet fixture should serialize"))
            .expect("worksheet fixture should parse");
    let (_budget, _cancellation, execution) = make_execution("worksheet-formula-snapshot");
    let resolver = Resolver::from_snapshot(&snapshot, SheetExtent::new(4, 5), &execution)
        .expect("snapshot worksheet index should build");

    assert!(matches!(
        read(&resolver, &execution, "Data", 1, 2),
        CellRead::Text(text) if text == "display"
    ));
}

#[test]
fn public_value_evaluate_reads_repeated_runs_arrays_and_finite_axes() {
    let sheets = vec![repeated_sheet()];
    let source_text = match &sheets[0].rows[0].cells[0].value {
        CellValue::Text(text) => text.as_str(),
        value => panic!("fixture text cell unexpectedly contains {value:?}"),
    };
    let source_ptr = source_text.as_ptr();
    let (_budget, _cancellation, execution) = make_execution("worksheet-formula-value-evaluate");
    let resolver = Resolver::new(&sheets, SheetExtent::new(4, 5), &execution)
        .expect("worksheet index should build");

    // Scalar projection consumes a singleton reference and retains the source
    // text borrow.  Both references land in the same physical repeated cell
    // run.
    for source in ["=[.A1]", "=[.C2]"] {
        let expression = Expression::parse(source)
            .unwrap_or_else(|error| panic!("{source:?} should parse: {error}"));
        let result = evaluate_worksheet(
            &expression,
            &resolver,
            &execution,
            "Data",
            0,
            0,
            Mode::Scalar,
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        let Value::Text(text) = result.value() else {
            panic!("{source:?} should return the stored text")
        };
        assert_eq!(text, source_text);
        assert!(ptr::eq(text.as_ptr(), source_ptr));
    }

    // The same compact number run is expanded only into the requested result
    // array.  The resolver still reads the two logical rows separately.
    let expression = Expression::parse("=0+[.D1:.D2]").expect("range should parse");
    let result = evaluate_worksheet(
        &expression,
        &resolver,
        &execution,
        "Data",
        0,
        0,
        Mode::Matrix,
    )
    .expect("range should evaluate through the worksheet resolver");
    let array = result.as_array().expect("range should produce an array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (2, 1));
    assert!(matches!(array.get(0), Some(Value::Number(value)) if value == 7.5));
    assert!(matches!(array.get(1), Some(Value::Number(value)) if value == 7.5));

    // Whole-axis references use the caller's finite extent and retain the
    // reference in matrix mode without reading every logically empty cell.
    let expression = Expression::parse("=[.$A:.$A]").expect("whole-column range should parse");
    let result = evaluate_worksheet(
        &expression,
        &resolver,
        &execution,
        "Data",
        0,
        0,
        Mode::Matrix,
    )
    .expect("whole-column range should resolve its finite extent");
    let Value::Reference(reference) = result.value() else {
        panic!("whole-column matrix result should retain a reference")
    };
    assert_eq!(reference.areas().len(), 1);
    assert_eq!(reference.areas()[0].starts(), [0, 0, 0]);
    assert_eq!(reference.areas()[0].extent(), [1, 4, 1]);
}

#[test]
fn public_value_evaluate_refuses_unsupported_cells_with_typed_errors() {
    let mut sheet = Sheet::new("Data").expect("sheet name is valid");
    let mut row = Row::new();
    let mut formula = Cell::new(CellValue::Number(99.0), "99");
    formula
        .set_formula("of:=99")
        .expect("formula cache is valid");
    row.push_cell(formula).expect("formula cell is valid");
    row.push_cell(Cell::new(CellValue::Date("2026-09-13".to_owned()), "date"))
        .expect("date cell is valid");
    row.push_cell(Cell::new(CellValue::Time("PT1H".to_owned()), "time"))
        .expect("time cell is valid");
    row.push_cell(Cell::new(
        CellValue::Unknown {
            kind: "future-type".to_owned(),
            value: Some("opaque".to_owned()),
        },
        "opaque",
    ))
    .expect("unknown cell is valid");
    sheet.push_row(row).expect("worksheet row is valid");
    let sheets = vec![sheet];
    let (_budget, _cancellation, execution) = make_execution("worksheet-formula-value-refusal");
    let resolver = Resolver::new(&sheets, SheetExtent::new(1, 4), &execution)
        .expect("worksheet index should build");

    for source in ["=+[.A1]", "=+[.B1]", "=+[.C1]", "=+[.D1]"] {
        let expression = Expression::parse(source)
            .unwrap_or_else(|error| panic!("{source:?} should parse: {error}"));
        let error = evaluate_worksheet(
            &expression,
            &resolver,
            &execution,
            "Data",
            0,
            0,
            Mode::Scalar,
        )
        .expect_err("formula and unsupported cells must refuse evaluation");
        assert!(matches!(
            error,
            EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
        ));
    }
}
