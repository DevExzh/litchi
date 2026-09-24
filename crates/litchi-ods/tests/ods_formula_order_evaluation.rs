//! Independent semantic coverage for the OpenFormula ordering/statistical
//! functions: `MEDIAN`, `MODE`, `LARGE`, `SMALL`, `PERCENTILE`,
//! `PERCENTRANK`, `QUARTILE`, and `RANK`.
//!
//! The fixture keeps references observable.  It has an ordered heterogeneous
//! Main sheet, two additional sheets for a 3-D reference, and two cells whose
//! values make a projected `MUNIT` parameter position-sensitive.

use std::{
    cell::{Cell, RefCell},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    SourceVersion,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Evaluated, Limits, Mode, OwnedValueView, Position, Resolver,
        SheetExtent, Value,
    },
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, ScalarValue,
    },
    expression::Expression,
};

const MAIN_NUMBERS: &str = "[.A1:.A6]";

#[derive(Debug, Clone)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
}

#[derive(Debug)]
struct OrderResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(String, usize, usize)>>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl OrderResolver {
    fn standard() -> Self {
        let mut resolver = Self {
            rows: 16,
            columns: 12,
            cells: (0..192).map(|_| FixtureCell::Empty).collect(),
            reads: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
            source_versions: None,
            source_version_calls: Cell::new(0),
        };

        for (row, value) in [1.0, 2.0, 2.0, 4.0, 7.0, 9.0].into_iter().enumerate() {
            resolver.set(row, 0, FixtureCell::Number(value));
        }
        for (row, value) in [10.0, 20.0, 30.0, 40.0].into_iter().enumerate() {
            resolver.set(row, 1, FixtureCell::Number(value));
        }
        resolver.set(0, 2, FixtureCell::Text("5".to_owned()));
        resolver.set(1, 2, FixtureCell::Empty);
        resolver.set(2, 2, FixtureCell::Logical(true));
        resolver.set(3, 2, FixtureCell::Error(ScalarError::NotAvailable));
        resolver.set(0, 6, FixtureCell::Number(1.0));
        resolver.set(1, 6, FixtureCell::Number(2.0));
        resolver
    }

    fn set(&mut self, row: usize, column: usize, value: FixtureCell) {
        let index = row
            .checked_mul(self.columns)
            .and_then(|index| index.checked_add(column))
            .expect("fixture coordinate fits");
        self.cells[index] = value;
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn clear_reads(&self) {
        self.reads.set(0);
        self.read_order.borrow_mut().clear();
    }
}

impl Resolver for OrderResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok((matches!(sheet, "Main" | "Data" | "Archive"))
            .then_some(SheetExtent::new(self.rows, self.columns)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.set(self.reads.get().saturating_add(1));
        self.read_order
            .borrow_mut()
            .push((sheet.to_owned(), row, column));
        if !matches!(sheet, "Main" | "Data" | "Archive") {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        if row >= self.rows || column >= self.columns {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        if sheet == "Data" && row == 0 && column == 0 {
            return Ok(CellRead::Number(100.0));
        }
        if sheet == "Archive" && row == 0 && column == 0 {
            return Ok(CellRead::Number(200.0));
        }
        if sheet != "Main" {
            return Ok(CellRead::Empty);
        }
        Ok(match &self.cells[row * self.columns + column] {
            FixtureCell::Empty => CellRead::Empty,
            FixtureCell::Number(value) => CellRead::Number(*value),
            FixtureCell::Logical(value) => CellRead::Logical(*value),
            FixtureCell::Text(value) => CellRead::Text(value.as_str()),
            FixtureCell::Error(error) => CellRead::Error(*error),
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        Ok(match sheet {
            "Main" => Some(0),
            "Data" => Some(1),
            "Archive" => Some(2),
            _ => None,
        })
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        Ok(match index {
            0 => Some("Main"),
            1 => Some("Data"),
            2 => Some("Archive"),
            _ => None,
        })
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(3)
    }

    fn source_version(
        &self,
        _execution: &ExecutionContext,
    ) -> Result<Option<SourceVersion>, EvaluationFailure> {
        let Some((expected, observed)) = self.source_versions else {
            return Ok(None);
        };
        let call = self.source_version_calls.get();
        self.source_version_calls.set(call.saturating_add(1));
        Ok(Some(if call == 0 { expected } else { observed }))
    }
}

fn execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one in-flight task"),
        NonZeroU64::new(1024 * 1024).expect("one in-flight MiB"),
        0,
    )
    .expect("valid execution limits");
    (
        budget.clone(),
        cancellation,
        ExecutionContext::new(budget, token, limits),
    )
}

fn parse(source: &str) -> Expression {
    Expression::parse(source).unwrap_or_else(|error| panic!("{source:?} should parse: {error}"))
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum OrderResult {
    Number(f64),
    Error(ScalarError),
    Other,
}

fn scalar_result(
    source: &str,
    execution: &ExecutionContext,
) -> Result<OrderResult, EvaluationFailure> {
    let expression = parse(source);
    let result = litchi_ods::codec::formula::evaluation::evaluate_scalar(
        &expression,
        &EvaluationContext::new(execution),
        &EvaluationLimits::default(),
    )?;
    Ok(match result.value() {
        ScalarValue::Number(value) => OrderResult::Number(*value),
        ScalarValue::Error(error) => OrderResult::Error(*error),
        _ => OrderResult::Other,
    })
}

fn value_result(
    source: &str,
    resolver: &OrderResolver,
    execution: &ExecutionContext,
    mode: Mode,
) -> Result<OrderResult, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, &Limits::default())?;
    Ok(match result.value() {
        Value::Number(value) => OrderResult::Number(value),
        Value::Error(error) => OrderResult::Error(error),
        _ => OrderResult::Other,
    })
}

fn evaluate<'a>(
    expression: &'a Expression,
    resolver: &'a OrderResolver,
    execution: &ExecutionContext,
    position: Position<'a>,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, position).with_mode(mode);
    value::evaluate(expression, resolver, &context, limits)
}

fn assert_close(result: OrderResult, expected: f64, source: &str) {
    let OrderResult::Number(actual) = result else {
        panic!("{source:?}: expected Number({expected}), got {result:?}");
    };
    let tolerance = expected.abs().max(actual.abs()).max(1.0) * 64.0 * f64::EPSILON;
    assert!(
        (actual - expected).abs() <= tolerance,
        "{source:?}: {actual} != {expected} (tolerance {tolerance})"
    );
}

fn assert_error(result: OrderResult, expected: ScalarError, source: &str) {
    assert_eq!(result, OrderResult::Error(expected), "{source:?}");
}

fn assert_array_numbers(result: &Evaluated<'_>, expected: &[f64], source: &str) {
    let array = result
        .as_array()
        .unwrap_or_else(|| panic!("{source:?}: expected an array result"));
    assert_eq!(array.len(), expected.len(), "{source:?} array length");
    for (index, expected) in expected.iter().copied().enumerate() {
        match array.get(index) {
            Some(Value::Number(actual)) => {
                assert_eq!(actual, expected, "{source:?}[{index}]")
            },
            other => panic!("{source:?}[{index}]: expected Number, got {other:?}"),
        }
    }
}

fn assert_array_shape_and_numbers(
    result: &Evaluated<'_>,
    rows: usize,
    columns: usize,
    expected: &[f64],
    source: &str,
) {
    let array = result
        .as_array()
        .unwrap_or_else(|| panic!("{source:?}: expected an array result"));
    assert_eq!(
        (array.shape().rows(), array.shape().columns()),
        (rows, columns),
        "{source:?} shape"
    );
    assert_eq!(array.len(), expected.len(), "{source:?} array length");
    for (index, expected) in expected.iter().copied().enumerate() {
        match array.get(index) {
            Some(Value::Number(actual)) => assert_close(
                OrderResult::Number(actual),
                expected,
                &format!("{source:?}[{index}]"),
            ),
            other => panic!("{source:?}[{index}]: expected Number, got {other:?}"),
        }
    }
}

#[test]
fn scalar_order_functions_cover_interpolation_ties_and_optional_parameters() {
    let (_budget, _cancellation, execution) = execution("ods-formula-order-scalar");
    for (source, expected) in [
        ("=MEDIAN(1;2;4;8)", 3.0),
        ("=MODE(1;2;2;4;4;7)", 2.0),
        // These functions have one sequence parameter followed by their
        // scalar parameter.  The scalar evaluator can therefore exercise
        // their direct profile with a one-member sequence; full
        // interpolation and ties are covered by the value arrays below.
        ("=LARGE(8;1)", 8.0),
        ("=SMALL(8;1)", 8.0),
        ("=PERCENTILE(8;0.5)", 8.0),
        ("=PERCENTRANK(8;8;3)", 1.0),
        ("=PERCENTRANK(8;8)", 1.0),
        ("=QUARTILE(8;2)", 8.0),
        ("=RANK(8;8)", 1.0),
        ("=RANK(8;8;1)", 1.0),
        ("=RANK(8;8;1.5)", 1.0),
    ] {
        assert_close(
            scalar_result(source, &execution)
                .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}")),
            expected,
            source,
        );
    }
}

#[test]
fn scalar_order_functions_reject_empty_invalid_rank_and_domain_inputs() {
    let (_budget, _cancellation, execution) = execution("ods-formula-order-invalid");
    for source in [
        "=MEDIAN()",
        "=MODE()",
        "=LARGE(1;0)",
        "=SMALL(1;2)",
        "=PERCENTILE(1;-0.1)",
        "=PERCENTILE(1;1.1)",
        "=PERCENTRANK(1;2)",
        "=PERCENTRANK(1;1;0)",
        "=PERCENTRANK(1;1;1.5)",
        "=QUARTILE(1;5)",
        "=RANK(2;1)",
    ] {
        assert_error(
            scalar_result(source, &execution).unwrap_or_else(|error| {
                panic!("{source:?} should produce a formula error: {error}")
            }),
            ScalarError::Value,
            source,
        );
    }
}

#[test]
fn scalar_order_functions_keep_formula_error_precedence() {
    let (_budget, _cancellation, execution) = execution("ods-formula-order-errors");
    for source in [
        "=MEDIAN(#N/A;1)",
        "=MODE(#N/A;1;1)",
        "=LARGE(#N/A;1)",
        "=SMALL(#N/A;1)",
        "=PERCENTILE(#N/A;0.5)",
        "=PERCENTRANK(#N/A;1;3)",
        "=QUARTILE(#N/A;2)",
        "=RANK(#N/A;1)",
    ] {
        assert_error(
            scalar_result(source, &execution).unwrap_or_else(|error| panic!("{source}: {error}")),
            ScalarError::NotAvailable,
            source,
        );
    }
}

#[test]
fn value_arrays_have_the_same_order_profile_in_scalar_and_matrix_modes() {
    let resolver = OrderResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-order-arrays");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (source, expected) in [
            ("=MEDIAN({1;2;4;8})", 3.0),
            ("=MODE({1;2;2|4;4;7})", 2.0),
            ("=LARGE({1;2;4;8};2)", 4.0),
            ("=SMALL({1;2;4;8};3)", 4.0),
            ("=PERCENTILE({1;2;4;8};0.25)", 1.75),
            ("=PERCENTRANK({1;2;4;8};3;3)", 0.5),
            ("=QUARTILE({1;2;4;8};1)", 1.75),
            ("=RANK(4;{1;2;4;8})", 2.0),
        ] {
            assert_close(
                value_result(source, &resolver, &execution, mode)
                    .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}")),
                expected,
                source,
            );
        }
    }
}

#[test]
fn large_and_small_accept_array_rank_parameters() {
    let resolver = OrderResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-order-rank-array");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (source, expected) in [
            ("=LARGE({1;2;4;8};{1;3})", &[8.0, 2.0][..]),
            ("=SMALL({1;2;4;8};{1;3})", &[1.0, 4.0][..]),
        ] {
            let expression = parse(source);
            let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(mode);
            let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
                .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}"));
            assert_array_numbers(&result, expected, source);
        }
    }
}

#[test]
fn large_and_small_array_ranks_pass_through_scalar_publication() {
    let resolver = OrderResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-order-scalar-array-ranks");
    for (source, rows, columns, expected) in [
        ("=LARGE({1;2;4;8};{1})", 1, 1, &[8.0][..]),
        ("=SMALL({1;2;4;8};{1})", 1, 1, &[1.0][..]),
        ("=LARGE({1;2;4;8};({1;3}))", 1, 2, &[8.0, 2.0][..]),
        ("=SMALL({1;2;4;8};({1;3}))", 1, 2, &[1.0, 4.0][..]),
        (
            "=IF(TRUE();LARGE({1;2;4;8};{1;2});0)",
            1,
            2,
            &[8.0, 4.0][..],
        ),
        (
            "=IF(TRUE();SMALL({1;2;4;8};{1;2});0)",
            1,
            2,
            &[1.0, 2.0][..],
        ),
    ] {
        let expression = parse(source);
        let result = evaluate(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Scalar,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source}: {error}"));
        assert_array_shape_and_numbers(&result, rows, columns, expected, source);
    }
}

#[test]
fn scalar_parameter_arrays_lift_order_functions_with_rectangular_broadcasting() {
    let resolver = OrderResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-order-parameter-arrays");

    // Number parameters are scalar in the scalar evaluator: a literal array
    // supplies its first element.  LARGE/SMALL are the explicit exception
    // covered by `large_and_small_accept_array_rank_parameters` above.
    for (source, expected) in [
        ("=PERCENTILE({1;2;4;8};{0;0.5})", 1.0),
        ("=PERCENTRANK({1;2;4;8};{2;6};3)", 0.333),
        ("=QUARTILE({1;2;4;8};{1;3})", 1.75),
        ("=RANK({2;4};{1;2;4;8};{0;1})", 3.0),
    ] {
        assert_close(
            value_result(source, &resolver, &execution, Mode::Scalar)
                .unwrap_or_else(|error| panic!("{source} in scalar mode: {error}")),
            expected,
            source,
        );
    }

    // Matrix mode lifts those same literal arrays and broadcasts them over
    // the complete argument shape.
    let mode = Mode::Matrix;
    let expression = parse("=PERCENTILE({1;2;4;8};{0;0.5})");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        mode,
        &Limits::default(),
    )
    .expect("PERCENTILE scalar-parameter array");
    assert_array_shape_and_numbers(&result, 1, 2, &[1.0, 3.0], "PERCENTILE parameter array");

    let expression = parse("=PERCENTRANK({1;2;4;8};{2;6};3)");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        mode,
        &Limits::default(),
    )
    .expect("PERCENTRANK scalar-parameter array");
    assert_array_shape_and_numbers(
        &result,
        1,
        2,
        &[0.333, 0.833],
        "PERCENTRANK parameter array",
    );

    let expression = parse("=QUARTILE({1;2;4;8};{1;3})");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        mode,
        &Limits::default(),
    )
    .expect("QUARTILE scalar-parameter array");
    assert_array_shape_and_numbers(&result, 1, 2, &[1.75, 5.0], "QUARTILE parameter array");

    // Sequence Data does not contribute output shape. The target and
    // order arrays therefore produce one result per parameter cell.
    let expression = parse("=RANK({2;4};{1;2;4;8};{0;1})");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        mode,
        &Limits::default(),
    )
    .expect("RANK scalar-parameter arrays");
    let array = result
        .as_array()
        .expect("RANK should retain parameter shape");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert!(matches!(array.get(0), Some(Value::Number(value)) if value == 3.0));
    assert!(matches!(array.get(1), Some(Value::Number(value)) if value == 3.0));

    // A column-shaped target/order pair keeps its two-cell shape; the
    // sequence Data argument still contributes no result dimensions.
    let expression = parse("=RANK({2|4};{1;2;4;8};{0|1})");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        mode,
        &Limits::default(),
    )
    .expect("RANK row/column broadcast");
    assert_array_shape_and_numbers(&result, 2, 1, &[3.0, 3.0], "RANK column parameters");

    // A row-shaped order array and a column-shaped target broadcast to a
    // 2x2 rectangle, independently of the sequence's four cells.
    let expression = parse("=RANK({2|4};{1;2;4;8};{0;1})");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        mode,
        &Limits::default(),
    )
    .expect("RANK rectangular broadcast");
    assert_array_shape_and_numbers(
        &result,
        2,
        2,
        &[3.0, 2.0, 2.0, 3.0],
        "RANK rectangular broadcast",
    );

    // MUNIT is an array-returning input to PERCENTILE. Matrix mode keeps its
    // complete identity parameter, while scalar mode consumes the [0,0]
    // parameter cell. The selected branch of a projected IF retains the
    // matrix parameter's output shape as well.
    let expression = parse("=PERCENTILE({1;2;4;8};MUNIT(2))");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("matrix PERCENTILE should retain MUNIT parameter");
    assert_array_shape_and_numbers(
        &result,
        2,
        2,
        &[8.0, 1.0, 1.0, 8.0],
        "matrix PERCENTILE MUNIT parameter",
    );

    assert_close(
        value_result(
            "=PERCENTILE({1;2;4;8};MUNIT(2))",
            &resolver,
            &execution,
            Mode::Scalar,
        )
        .expect("scalar PERCENTILE should project MUNIT parameter"),
        8.0,
        "scalar PERCENTILE MUNIT parameter",
    );

    let expression = parse("=IF({TRUE()|TRUE()};PERCENTILE({1;2;4;8};MUNIT(2));0)");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("projected IF should retain MUNIT parameter output shape");
    assert_array_shape_and_numbers(
        &result,
        2,
        2,
        &[8.0, 1.0, 1.0, 8.0],
        "projected IF PERCENTILE MUNIT parameter",
    );
}

#[test]
fn references_lists_and_three_dimensional_ranges_follow_sequence_signatures() {
    let resolver = OrderResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-order-references");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (source, expected) in [
            (format!("=MEDIAN({MAIN_NUMBERS})"), 3.0),
            (format!("=MODE({MAIN_NUMBERS})"), 2.0),
            (format!("=LARGE({MAIN_NUMBERS};2)"), 7.0),
            (format!("=SMALL({MAIN_NUMBERS};3)"), 2.0),
            (format!("=PERCENTILE({MAIN_NUMBERS};0.25)"), 2.0),
            (format!("=PERCENTRANK({MAIN_NUMBERS};4)"), 0.6),
            (format!("=QUARTILE({MAIN_NUMBERS};3)"), 6.25),
            (format!("=RANK(4;{MAIN_NUMBERS})"), 3.0),
        ] {
            let result = value_result(&source, &resolver, &execution, mode)
                .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}"));
            assert_close(result, expected, &source);
        }
    }

    for (source, expected) in [
        ("=MEDIAN([.A1:.A3]~[.B1:.B3])", 6.0),
        ("=LARGE([.A1:.A3]~[.B1:.B3];2)", 20.0),
        ("=SMALL([.A1:.A3]~[.B1:.B3];2)", 2.0),
        ("=PERCENTILE([.A1:.A3]~[.B1:.B3];0.5)", 6.0),
        ("=PERCENTRANK([.A1:.A3]~[.B1:.B3];10)", 0.6),
        ("=RANK(10;[.A1:.A3]~[.B1:.B3])", 3.0),
    ] {
        assert_close(
            value_result(source, &resolver, &execution, Mode::Matrix)
                .unwrap_or_else(|error| panic!("{source}: {error}")),
            expected,
            source,
        );
    }

    for (source, expected) in [
        ("=MEDIAN([Main.A1:Archive.A1])", 100.0),
        ("=LARGE([Main.A1:Archive.A1];2)", 100.0),
        ("=SMALL([Main.A1:Archive.A1];2)", 100.0),
        ("=PERCENTILE([Main.A1:Archive.A1];0.5)", 100.0),
        ("=PERCENTRANK([Main.A1:Archive.A1];100)", 0.5),
        ("=QUARTILE([Main.A1:Archive.A1];2)", 100.0),
        ("=RANK(100;[Main.A1:Archive.A1])", 2.0),
    ] {
        assert_close(
            value_result(source, &resolver, &execution, Mode::Matrix)
                .unwrap_or_else(|error| panic!("{source}: {error}")),
            expected,
            source,
        );
    }
}

#[test]
fn number_sequence_reducers_retain_shape_refusal_for_reference_lists() {
    let resolver = OrderResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-order-shapes");
    for function in ["MODE", "QUARTILE"] {
        let source = if function == "QUARTILE" {
            format!("={function}([.A1:.A2]~[.B1:.B2];1)")
        } else {
            format!("={function}([.A1:.A2]~[.B1:.B2])")
        };
        resolver.clear_reads();
        assert_error(
            value_result(&source, &resolver, &execution, Mode::Matrix)
                .unwrap_or_else(|error| panic!("{source}: {error}")),
            ScalarError::Value,
            &source,
        );
        assert_eq!(resolver.reads(), 0, "{source} must refuse before reads");
    }
}

#[test]
fn scalar_parameter_reference_lists_refuse_before_any_data_read() {
    let resolver = OrderResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-order-parameter-lists");
    for source in [
        "=RANK([.A1:.A2]~[.A3:.A4];[.B1:.B6])",
        "=PERCENTILE([.B1:.B6];[.A1:.A2]~[.A3:.A4])",
    ] {
        resolver.clear_reads();
        assert_error(
            value_result(source, &resolver, &execution, Mode::Matrix)
                .unwrap_or_else(|error| panic!("{source}: {error}")),
            ScalarError::Value,
            source,
        );
        assert_eq!(resolver.reads(), 0, "{source} must refuse before any reads");
    }
}

#[test]
fn reference_errors_win_over_empty_or_generated_order_results() {
    let resolver = OrderResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-order-reference-errors");
    for source in [
        "=MEDIAN([.C4])",
        "=MODE([.C4];1)",
        "=LARGE([.C4];1)",
        "=SMALL([.C4];1)",
        "=PERCENTILE([.C4];0.5)",
        "=PERCENTRANK([.C4];0)",
        "=QUARTILE([.C4];1)",
        "=RANK(1;[.C4])",
    ] {
        assert_error(
            value_result(source, &resolver, &execution, Mode::Matrix).unwrap(),
            ScalarError::NotAvailable,
            source,
        );
    }
}

#[test]
fn projected_order_reducers_cache_complete_references_but_recompute_munit_parameters() {
    let mut resolver = OrderResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-order-projection");

    let expression = parse("=IF({TRUE()|TRUE()};MEDIAN([.A1:.A6]);0)");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("projected median");
    assert_array_numbers(&result, &[3.0, 3.0], "projected median");
    assert_eq!(
        resolver.reads(),
        6,
        "complete reference should be scanned once"
    );

    let expression =
        parse("=IF({TRUE()|TRUE()};PERCENTILE([.A1:.A6];COUNT(MUNIT([.G1:.G2]))/10);0)");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("position-sensitive MUNIT percentile");
    // MUNIT's matrix argument is scalar-converted from its first parameter
    // cell in this direct reducer context. Nested conditional criteria below
    // retain their separate position-sensitive scalar path.
    assert_array_numbers(&result, &[1.5, 1.5], "MUNIT percentile projection");

    // Make the second-largest value a duplicate so the conditional criterion
    // exposes both COUNT(MUNIT(...)) sizes: k=1 matches once, k=4 matches
    // twice after the data change.
    resolver.set(3, 0, FixtureCell::Number(7.0));
    resolver.clear_reads();
    let expression =
        parse("=IF({TRUE()|TRUE()};COUNTIF([.A1:.A6];LARGE([.A1:.A6];COUNT(MUNIT([.G1:.G2]))));0)");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("position-sensitive nested order criterion");
    assert_array_numbers(&result, &[1.0, 2.0], "nested order criterion");
    assert!(
        resolver.reads() >= 6,
        "criterion must inspect the complete range"
    );

    // A bare reference in a scalar parameter position remains position
    // sensitive under a projected conditional.  The surrounding reducer can
    // cache its complete data reference, while LARGE must consume G1 and G2
    // separately for the two projected cells.
    resolver.clear_reads();
    let expression = parse("=IF({TRUE()|TRUE()};COUNTIF([.A1:.A6];LARGE([.A1:.A6];[.G1:.G2]));0)");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("position-sensitive scalar-reference criterion");
    assert_array_shape_and_numbers(
        &result,
        2,
        1,
        &[1.0, 2.0],
        "scalar-reference order criterion",
    );
}

#[test]
fn projected_order_scalar_parameters_lift_references_and_computed_arrays() {
    let mut resolver = OrderResolver::standard();
    resolver.set(0, 6, FixtureCell::Number(0.1));
    resolver.set(1, 6, FixtureCell::Number(0.2));
    let (_budget, _cancellation, execution) = execution("ods-formula-order-parameter-shapes");

    for source in [
        "=IF({TRUE()|TRUE()};PERCENTILE([.A1:.A6];[.G1:.G2]);0)",
        "=IF({TRUE()|TRUE()};PERCENTILE([.A1:.A6];TRANSPOSE({0.1;0.2}));0)",
    ] {
        let expression = parse(source);
        let result = evaluate(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source}: {error}"));
        assert_array_shape_and_numbers(&result, 2, 1, &[1.5, 2.0], source);
    }
}

#[test]
fn top_level_matrix_order_parameters_retain_reference_and_computed_shapes() {
    let mut resolver = OrderResolver::standard();
    resolver.set(0, 6, FixtureCell::Number(0.1));
    resolver.set(1, 6, FixtureCell::Number(0.2));
    resolver.set(0, 8, FixtureCell::Number(0.0));
    resolver.set(1, 8, FixtureCell::Number(1.0));
    let (_budget, _cancellation, execution) = execution("ods-formula-order-top-level-shapes");

    let expression = parse("=PERCENTILE([.A1:.A6];[.G1:.G2])");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("top-level PERCENTILE reference parameter");
    assert_array_shape_and_numbers(
        &result,
        2,
        1,
        &[1.5, 2.0],
        "top-level PERCENTILE reference parameter",
    );

    let expression = parse("=RANK([.G1:.G2];[.A1:.A6];[.I1:.I2])");
    resolver.set(0, 6, FixtureCell::Number(2.0));
    resolver.set(1, 6, FixtureCell::Number(4.0));
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("top-level RANK reference parameters");
    assert_array_shape_and_numbers(
        &result,
        2,
        1,
        &[4.0, 4.0],
        "top-level RANK reference parameters",
    );

    resolver.set(0, 6, FixtureCell::Number(0.1));
    resolver.set(1, 6, FixtureCell::Number(0.2));
    for source in [
        "=PERCENTILE([.A1:.A6];{0.1|0.2}+0)",
        "=PERCENTILE([.A1:.A6];TRANSPOSE(TRANSPOSE({0.1|0.2})))",
    ] {
        let expression = parse(source);
        let result = evaluate(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source}: {error}"));
        assert_array_shape_and_numbers(&result, 2, 1, &[1.5, 2.0], source);
    }
}

#[test]
fn scalar_if_condition_does_not_hide_order_parameter_reference_shape() {
    let mut resolver = OrderResolver::standard();
    resolver.set(0, 6, FixtureCell::Number(0.1));
    resolver.set(1, 6, FixtureCell::Number(0.2));
    let (_budget, _cancellation, execution) = execution("ods-formula-order-scalar-if-shape");
    let expression = parse("=IF(TRUE();PERCENTILE([.A1:.A6];[.G1:.G2]);0)");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("scalar IF should retain the selected parameter shape");
    assert_array_shape_and_numbers(
        &result,
        2,
        1,
        &[1.5, 2.0],
        "scalar IF PERCENTILE parameter shape",
    );
}

#[test]
fn nested_sum_consumes_complete_order_argument_per_projected_cell() {
    let mut resolver = OrderResolver::standard();
    resolver.set(0, 6, FixtureCell::Number(0.1));
    resolver.set(1, 6, FixtureCell::Number(0.2));
    let (_budget, _cancellation, execution) = execution("ods-formula-order-nested-sum-cache");
    let expression = parse("=IF({TRUE()|TRUE()};SUM(PERCENTILE([.A1:.A6];[.G1:.G2]));0)");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("nested SUM should consume the complete order result");
    assert_array_shape_and_numbers(
        &result,
        2,
        1,
        &[3.5, 3.5],
        "nested SUM complete order result",
    );
}

#[test]
fn order_results_can_be_owned_without_retaining_reference_state() {
    let resolver = OrderResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-order-owned");
    let expression = parse("=MEDIAN([.A1:.A6])");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("median reference result");
    let owned = result
        .to_owned(&execution, &Limits::default())
        .expect("order result should be ownable");
    assert!(matches!(owned.value(), OwnedValueView::Number(value) if value == 3.0));
}
