//! Independent semantic coverage for the OpenFormula paired-statistics and
//! regression functions: `CORREL`, `COVAR`, `PEARSON`, `RSQ`, `SLOPE`,
//! `INTERCEPT`, `STEYX`, and `FORECAST`.
//!
//! The fixture keeps the two data arguments observable and rectangular.  It
//! distinguishes complete ForceArray references from scalar FORECAST queries,
//! records read order, and leaves the source/version hooks available for the
//! typed-failure and projected-branch cases.

use std::{
    cell::{Cell, RefCell},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Evaluated, Limits, Mode, Position, Resolver, SheetExtent, Value,
    },
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, ScalarValue,
    },
    expression::Expression,
};

#[derive(Debug, Clone)]
enum FixtureCell {
    Empty,
    Number(f64),
    Text(String),
    Error(ScalarError),
}

#[derive(Debug)]
struct PairedResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(String, usize, usize)>>,
}

impl PairedResolver {
    fn standard() -> Self {
        let mut resolver = Self {
            rows: 16,
            columns: 12,
            cells: (0..192).map(|_| FixtureCell::Empty).collect(),
            reads: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
        };

        // The primary fit is y = 2x + 1.  The same values make the
        // population covariance, correlation, slope, intercept, STEYX, and
        // FORECAST expectations exact or very small to tolerance.
        for (row, (x, y)) in [(1.0, 3.0), (2.0, 5.0), (3.0, 7.0), (4.0, 9.0)]
            .into_iter()
            .enumerate()
        {
            resolver.set(row, 0, FixtureCell::Number(x));
            resolver.set(row, 1, FixtureCell::Number(y));
        }

        // C has one omitted Text member and one omitted Empty member.  The
        // aligned A member is omitted with it, rather than being converted.
        resolver.set(0, 2, FixtureCell::Number(3.0));
        resolver.set(1, 2, FixtureCell::Text("ignored".to_owned()));
        resolver.set(2, 2, FixtureCell::Number(7.0));
        resolver.set(3, 2, FixtureCell::Empty);

        // E is a criterion target with duplicate fitted values, so a nested
        // FORECAST query exposes position-sensitive MUNIT sizes.
        for (row, value) in [3.0, 3.0, 3.0, 9.0].into_iter().enumerate() {
            resolver.set(row, 4, FixtureCell::Number(value));
        }

        // G1:G2 are scalar FORECAST queries.  The matrix query tests widen to
        // this column shape; the conditional MUNIT case consumes them at the
        // criterion position.
        resolver.set(0, 6, FixtureCell::Number(5.0));
        resolver.set(1, 6, FixtureCell::Number(6.0));
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

impl Resolver for PairedResolver {
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
        if !matches!(sheet, "Main" | "Data" | "Archive")
            || row >= self.rows
            || column >= self.columns
        {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        if sheet != "Main" {
            return Ok(CellRead::Empty);
        }
        Ok(match &self.cells[row * self.columns + column] {
            FixtureCell::Empty => CellRead::Empty,
            FixtureCell::Number(value) => CellRead::Number(*value),
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
enum ResultValue {
    Number(f64),
    Error(ScalarError),
    Other,
}

fn scalar_result(
    source: &str,
    execution: &ExecutionContext,
) -> Result<ResultValue, EvaluationFailure> {
    let expression = parse(source);
    let result = litchi_ods::codec::formula::evaluation::evaluate_scalar(
        &expression,
        &EvaluationContext::new(execution),
        &EvaluationLimits::default(),
    )?;
    Ok(match result.value() {
        ScalarValue::Number(value) => ResultValue::Number(*value),
        ScalarValue::Error(error) => ResultValue::Error(*error),
        _ => ResultValue::Other,
    })
}

fn value_result(
    source: &str,
    resolver: &PairedResolver,
    execution: &ExecutionContext,
    mode: Mode,
) -> Result<ResultValue, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, &Limits::default())?;
    Ok(match result.value() {
        Value::Number(value) => ResultValue::Number(value),
        Value::Error(error) => ResultValue::Error(error),
        _ => ResultValue::Other,
    })
}

fn evaluate<'a>(
    expression: &'a Expression,
    resolver: &'a PairedResolver,
    execution: &ExecutionContext,
    position: Position<'a>,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, position).with_mode(mode);
    value::evaluate(expression, resolver, &context, limits)
}

fn assert_close(result: ResultValue, expected: f64, source: &str) {
    let ResultValue::Number(actual) = result else {
        panic!("{source:?}: expected Number({expected}), got {result:?}");
    };
    let tolerance = expected.abs().max(actual.abs()).max(1.0) * 64.0 * f64::EPSILON;
    assert!(
        (actual - expected).abs() <= tolerance,
        "{source:?}: {actual} != {expected} (tolerance {tolerance})"
    );
}

fn assert_error(result: ResultValue, expected: ScalarError, source: &str) {
    assert_eq!(result, ResultValue::Error(expected), "{source:?}");
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
    assert_eq!(array.len(), expected.len(), "{source:?} length");
    for (index, expected) in expected.iter().copied().enumerate() {
        match array.get(index) {
            Some(Value::Number(actual)) => assert_close(
                ResultValue::Number(actual),
                expected,
                &format!("{source:?}[{index}]"),
            ),
            other => panic!("{source:?}[{index}]: expected Number, got {other:?}"),
        }
    }
}

#[test]
fn scalar_paired_functions_cover_direct_one_by_one_and_regression_domains() {
    let (_budget, _cancellation, execution) = execution("ods-formula-paired-scalar");

    // A scalar is the one-by-one ForceArray instance. COVAR is defined for a
    // single admitted pair; the variance-based functions expose their
    // selected zero-variance subtype.
    assert_close(
        scalar_result("=COVAR(1;2)", &execution).expect("scalar COVAR"),
        0.0,
        "=COVAR(1;2)",
    );
    for source in [
        "=CORREL(1;2)",
        "=PEARSON(1;2)",
        "=RSQ(1;2)",
        "=SLOPE(2;1)",
        "=INTERCEPT(2;1)",
        "=FORECAST(5;2;1)",
    ] {
        assert_error(
            scalar_result(source, &execution).expect("scalar paired formula value"),
            ScalarError::DivisionByZero,
            source,
        );
    }
    assert_error(
        scalar_result("=STEYX(2;1)", &execution).expect("scalar STEYX formula value"),
        ScalarError::Value,
        "=STEYX(2;1)",
    );
}

#[test]
fn scalar_paired_functions_reject_arity_and_keep_formula_errors() {
    let (_budget, _cancellation, execution) = execution("ods-formula-paired-scalar-errors");
    for source in [
        "=CORREL()",
        "=COVAR(1)",
        "=PEARSON(1;2;3)",
        "=RSQ(1)",
        "=SLOPE(1)",
        "=INTERCEPT(1;2;3)",
        "=STEYX(1;2)",
        "=FORECAST(1;2)",
        "=CORREL(1;)",
        "=FORECAST(;1;2)",
    ] {
        assert_error(
            scalar_result(source, &execution).expect("arity is a formula value"),
            ScalarError::Value,
            source,
        );
    }
    for (source, expected) in [
        ("=CORREL(#N/A;1)", ScalarError::NotAvailable),
        ("=COVAR(1;#DIV/0!)", ScalarError::DivisionByZero),
        ("=PEARSON(#N/A;1)", ScalarError::NotAvailable),
        ("=RSQ(1;#N/A)", ScalarError::NotAvailable),
        ("=SLOPE(#N/A;1)", ScalarError::NotAvailable),
        ("=INTERCEPT(1;#N/A)", ScalarError::NotAvailable),
        ("=STEYX(#N/A;1)", ScalarError::NotAvailable),
        ("=FORECAST(1;#N/A;1)", ScalarError::NotAvailable),
    ] {
        assert_error(
            scalar_result(source, &execution).expect("formula error is retained"),
            expected,
            source,
        );
    }
}

#[test]
fn forecast_query_conversion_errors_match_scalar_and_value_precedence() {
    let resolver = PairedResolver::standard();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-paired-query-conversion-precedence");
    for source in [
        "=FORECAST(\"bad\";#N/A;1)",
        "=FORECAST(;#N/A;1)",
        "=FORECAST(\"1e999\";#N/A;1)",
    ] {
        assert_error(
            scalar_result(source, &execution).expect("scalar query conversion result"),
            ScalarError::NotAvailable,
            source,
        );
        for mode in [Mode::Scalar, Mode::Matrix] {
            assert_error(
                value_result(source, &resolver, &execution, mode)
                    .expect("value query conversion result"),
                ScalarError::NotAvailable,
                source,
            );
        }
    }

    let source = "=FORECAST(#DIV/0!;#N/A;1)";
    assert_error(
        scalar_result(source, &execution).expect("scalar query formula error result"),
        ScalarError::DivisionByZero,
        source,
    );
    for mode in [Mode::Scalar, Mode::Matrix] {
        assert_error(
            value_result(source, &resolver, &execution, mode)
                .expect("value query formula error result"),
            ScalarError::DivisionByZero,
            source,
        );
    }
}

#[test]
fn paired_arrays_have_the_same_profile_in_scalar_and_matrix_modes() {
    let resolver = PairedResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-paired-arrays");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (source, expected) in [
            ("=CORREL({1;2;3;4};{3;5;7;9})", 1.0),
            ("=COVAR({1;2;3;4};{3;5;7;9})", 2.5),
            ("=PEARSON({1;2;3;4};{3;5;7;9})", 1.0),
            ("=RSQ({3;5;7;9};{1;2;3;4})", 1.0),
            ("=SLOPE({3;5;7;9};{1;2;3;4})", 2.0),
            ("=INTERCEPT({3;5;7;9};{1;2;3;4})", 1.0),
            ("=STEYX({3;5;7;9};{1;2;3;4})", 0.0),
            ("=FORECAST(5;{3;5;7;9};{1;2;3;4})", 11.0),
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
fn paired_equations_cover_nonperfect_fit_and_pair_omission() {
    let resolver = PairedResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-paired-equations");
    for (source, expected) in [
        ("=COVAR({1;2;3;4};{2;4;5;7})", 2.0),
        ("=CORREL({1;2;3;4};{2;4;5;7})", 8.0 / 65.0_f64.sqrt()),
        ("=PEARSON({1;2;3;4};{2;4;5;7})", 8.0 / 65.0_f64.sqrt()),
        ("=RSQ({2;4;5;7};{1;2;3;4})", 64.0 / 65.0),
        ("=SLOPE({2;4;5;7};{1;2;3;4})", 1.6),
        ("=INTERCEPT({2;4;5;7};{1;2;3;4})", 0.5),
        ("=STEYX({2;4;5;7};{1;2;3;4})", 0.1_f64.sqrt()),
        ("=FORECAST(5;{2;4;5;7};{1;2;3;4})", 8.5),
        ("=SLOPE({3;\"ignored\";7;9};{1;2;3;4})", 2.0),
        ("=FORECAST(5;{3;TRUE();7;9};{1;2;3;4})", 11.0),
    ] {
        assert_close(
            value_result(source, &resolver, &execution, Mode::Matrix)
                .unwrap_or_else(|error| panic!("{source}: {error}")),
            expected,
            source,
        );
    }
    assert_error(
        value_result(
            "=RSQ({\"text\";TRUE()};{\"other\";FALSE()})",
            &resolver,
            &execution,
            Mode::Matrix,
        )
        .expect("RSQ empty valid shape"),
        ScalarError::NotAvailable,
        "RSQ empty valid shape",
    );
    assert_eq!(
        resolver.reads(),
        0,
        "literal arrays do not read the resolver"
    );
}

#[test]
fn rectangular_references_are_paired_in_both_modes_and_3d_is_rejected() {
    let resolver = PairedResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-paired-references");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (source, expected) in [
            ("=CORREL([.A1:.A4];[.B1:.B4])", 1.0),
            ("=COVAR([.A1:.A4];[.B1:.B4])", 2.5),
            ("=PEARSON([.A1:.A4];[.B1:.B4])", 1.0),
            ("=RSQ([.B1:.B4];[.A1:.A4])", 1.0),
            ("=SLOPE([.B1:.B4];[.A1:.A4])", 2.0),
            ("=INTERCEPT([.B1:.B4];[.A1:.A4])", 1.0),
            ("=STEYX([.B1:.B4];[.A1:.A4])", 0.0),
            ("=FORECAST(6;[.B1:.B4];[.A1:.A4])", 13.0),
        ] {
            assert_close(
                value_result(source, &resolver, &execution, mode)
                    .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}")),
                expected,
                source,
            );
        }
    }
    assert_eq!(
        resolver.reads(),
        128,
        "each two-column fit is read in both modes"
    );

    resolver.clear_reads();
    for source in [
        "=CORREL([.A1:.A2]~[.A3:.A4];[.B1:.B4])",
        "=COVAR([.A1:.A2]~[.A3:.A4];[.B1:.B4])",
        "=PEARSON([.A1:.A2]~[.A3:.A4];[.B1:.B4])",
        "=RSQ([.A1:.A2]~[.A3:.A4];[.B1:.B4])",
        "=SLOPE([.B1:.B4];[.A1:.A2]~[.A3:.A4])",
        "=INTERCEPT([.B1:.B4];[.A1:.A2]~[.A3:.A4])",
        "=STEYX([.B1:.B4];[.A1:.A2]~[.A3:.A4])",
        "=FORECAST(5;[.B1:.B4];[.A1:.A2]~[.A3:.A4])",
        "=CORREL([Main.A1:Archive.A2];[Main.A1:Archive.A2])",
    ] {
        assert_error(
            value_result(source, &resolver, &execution, Mode::Matrix)
                .unwrap_or_else(|error| panic!("{source}: {error}")),
            ScalarError::Value,
            source,
        );
    }
    assert_eq!(
        resolver.reads(),
        0,
        "list and 3-D shape refusals are read-free"
    );
}

#[test]
fn paired_shape_errors_distinguish_rsquared_count_from_geometry() {
    let resolver = PairedResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-paired-shapes");
    for function in ["CORREL", "COVAR", "PEARSON", "SLOPE", "INTERCEPT", "STEYX"] {
        let source = format!("={function}({{1;2;3}};{{1;2}})");
        assert_error(
            value_result(&source, &resolver, &execution, Mode::Matrix).unwrap(),
            ScalarError::Value,
            &source,
        );
        let source = format!("={function}({{1;2;3}};{{1|2|3}})");
        assert_error(
            value_result(&source, &resolver, &execution, Mode::Matrix).unwrap(),
            ScalarError::Value,
            &source,
        );
    }
    assert_error(
        value_result(
            "=FORECAST(1;{1;2;3};{1;2})",
            &resolver,
            &execution,
            Mode::Matrix,
        )
        .unwrap(),
        ScalarError::Value,
        "FORECAST shape",
    );
    assert_error(
        value_result("=RSQ({1;2;3};{1;2})", &resolver, &execution, Mode::Matrix).unwrap(),
        ScalarError::NotAvailable,
        "RSQ different total cell count",
    );
    assert_error(
        value_result("=RSQ({1;2;3};{1|2|3})", &resolver, &execution, Mode::Matrix).unwrap(),
        ScalarError::Value,
        "RSQ equal count different geometry",
    );
    assert_eq!(resolver.reads(), 0, "literal shape gates do not read");
}

#[test]
fn paired_reference_errors_are_retained_in_source_argument_order() {
    let mut resolver = PairedResolver::standard();
    // The first argument has its error in a later cell; the second has an
    // earlier error. The paired reducer's source argument order still makes
    // the first argument's formula error win.
    resolver.set(3, 0, FixtureCell::Error(ScalarError::NotAvailable));
    resolver.set(0, 1, FixtureCell::Error(ScalarError::DivisionByZero));
    let (_budget, _cancellation, execution) = execution("ods-formula-paired-reference-errors");
    assert_error(
        value_result(
            "=CORREL([.A1:.A4];[.B1:.B4])",
            &resolver,
            &execution,
            Mode::Matrix,
        )
        .unwrap(),
        ScalarError::NotAvailable,
        "first argument formula error precedence",
    );
    assert_eq!(resolver.reads(), 8, "the admitted scan completes");
}

#[test]
fn source_formula_errors_beat_generated_nonfinite_numbers() {
    let (_budget, _cancellation, execution) =
        execution("ods-formula-paired-generated-error-precedence");

    let mut resolver = PairedResolver::standard();
    resolver.set(0, 0, FixtureCell::Number(f64::INFINITY));
    resolver.set(0, 1, FixtureCell::Error(ScalarError::NotAvailable));
    assert_error(
        value_result(
            "=CORREL([.A1:.A4];[.B1:.B4])",
            &resolver,
            &execution,
            Mode::Matrix,
        )
        .expect("paired nonfinite/formula scan"),
        ScalarError::NotAvailable,
        "CORREL source formula error beats generated Number error",
    );

    let mut resolver = PairedResolver::standard();
    resolver.set(0, 6, FixtureCell::Number(f64::INFINITY));
    resolver.set(0, 1, FixtureCell::Error(ScalarError::NotAvailable));
    assert_error(
        value_result(
            "=FORECAST([.G1];[.B1:.B4];[.A1:.A4])",
            &resolver,
            &execution,
            Mode::Matrix,
        )
        .expect("FORECAST nonfinite query/formula scan"),
        ScalarError::NotAvailable,
        "FORECAST source formula error beats generated query Number error",
    );
}

#[test]
fn forecast_query_lifts_references_and_computed_arrays_but_scalar_mode_projects() {
    let resolver = PairedResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-paired-query-shapes");

    let expression = parse("=FORECAST({5;6};{3;5;7;9};{1;2;3;4})");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("matrix FORECAST query array");
    assert_array_shape_and_numbers(&result, 1, 2, &[11.0, 13.0], "FORECAST query array");

    assert_close(
        value_result(
            "=FORECAST({5;6};{3;5;7;9};{1;2;3;4})",
            &resolver,
            &execution,
            Mode::Scalar,
        )
        .expect("scalar FORECAST query projection"),
        11.0,
        "scalar FORECAST query projection",
    );

    let projected_queries = [
        (
            "=IF(TRUE();FORECAST([.G1:.G2];[.B1:.B4];[.A1:.A4]);0)",
            2,
            1,
        ),
        (
            "=IF(TRUE();FORECAST(TRANSPOSE({5|6});[.B1:.B4];[.A1:.A4]);0)",
            1,
            2,
        ),
    ];
    for (source, rows, columns) in projected_queries {
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
        assert_array_shape_and_numbers(&result, rows, columns, &[11.0, 13.0], source);
    }
}

#[test]
fn forecast_cached_fit_preserves_zero_variance_error_across_projected_queries() {
    let resolver = PairedResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-paired-zero-variance-cache");
    let expression = parse("=IF(TRUE();FORECAST([.G1:.G2];[.B1:.B4];{1|1|1|1});0)");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("projected FORECAST zero-variance result");
    let array = result
        .as_array()
        .expect("projected FORECAST result is an array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (2, 1));
    for index in 0..array.len() {
        assert_eq!(
            array.get(index),
            Some(Value::Error(ScalarError::DivisionByZero)),
            "FORECAST projected error at {index}"
        );
    }
    assert_eq!(
        resolver.reads(),
        6,
        "the invariant fit is scanned once before both query coordinates"
    );
}

#[test]
fn forecast_consumes_munit_output_as_an_array_and_keeps_nested_criterion_position() {
    let mut resolver = PairedResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-paired-munit");

    let expression = parse("=FORECAST(MUNIT(2);{3;5;7;9};{1;2;3;4})");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("FORECAST MUNIT output");
    assert_array_shape_and_numbers(&result, 2, 2, &[3.0, 1.0, 1.0, 3.0], "FORECAST MUNIT");

    resolver.set(0, 6, FixtureCell::Number(1.0));
    resolver.set(1, 6, FixtureCell::Number(2.0));
    let expression = parse(
        "=IF({TRUE()|TRUE()};COUNTIF([.E1:.E4];FORECAST(COUNT(MUNIT([.G1:.G2]));[.B1:.B4];[.A1:.A4]));0)",
    );
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("nested FORECAST criterion");
    assert_array_shape_and_numbers(
        &result,
        2,
        1,
        &[3.0, 1.0],
        "nested FORECAST MUNIT criterion",
    );
}
