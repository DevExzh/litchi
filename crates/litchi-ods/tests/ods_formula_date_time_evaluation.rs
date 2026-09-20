//! Semantic and scalar/value differential coverage for the ODF 1.4 date/time
//! family.  The timestamp is always injected explicitly; this target never
//! consults the host clock.

use std::{
    cell::{Cell, RefCell},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Limits, Mode, Position, Resolver, SheetExtent, Value,
    },
    evaluation::{
        CalculationTimestamp, EvaluationContext, EvaluationFailure, EvaluationLimits,
        EvaluationOptions, ScalarError, ScalarValue, UnsupportedKind, evaluate_scalar,
    },
    expression::Expression,
};

const TIMESTAMP_SERIAL: f64 = 46_000.5;

#[derive(Debug, Clone)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
}

#[derive(Debug)]
struct DateResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(String, usize, usize)>>,
}

impl DateResolver {
    fn standard() -> Self {
        let rows = 32;
        let columns = 16;
        let mut resolver = Self {
            rows,
            columns,
            cells: vec![FixtureCell::Empty; rows * columns],
            reads: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
        };

        // Holiday candidates: Jan 1 and Jan 8, followed by Empty, Text, and
        // a formula Error so sequence scanning and precedence are observable.
        resolver.set(0, 0, FixtureCell::Number(45_292.0));
        resolver.set(1, 0, FixtureCell::Number(45_299.0));
        resolver.set(2, 0, FixtureCell::Empty);
        resolver.set(3, 0, FixtureCell::Text("skip".to_owned()));
        resolver.set(4, 0, FixtureCell::Error(ScalarError::NotAvailable));
        resolver.set(5, 0, FixtureCell::Number(45_300.0));

        // Sunday-through-Saturday custom workweek: Sunday and Saturday are
        // non-workdays, the five middle entries are workdays.
        for (row, value) in [true, false, false, false, false, false, true]
            .into_iter()
            .enumerate()
        {
            resolver.set(row, 1, FixtureCell::Logical(value));
        }

        // MUNIT probes use two position-sensitive scalar sizes.
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

    fn read_order(&self) -> Vec<(String, usize, usize)> {
        self.read_order.borrow().clone()
    }
}

impl Resolver for DateResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(SheetExtent::new(self.rows, self.columns)))
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
        if sheet != "Main" || row >= self.rows || column >= self.columns {
            return Ok(CellRead::Error(ScalarError::Reference));
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
        Ok((sheet == "Main").then_some(0))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        Ok((index == 0).then_some("Main"))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(1)
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

fn options() -> EvaluationOptions {
    EvaluationOptions::default().with_calculation_timestamp(
        CalculationTimestamp::from_serial(TIMESTAMP_SERIAL)
            .expect("test timestamp is in the date profile"),
    )
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum Observed {
    Number(f64),
    Error(ScalarError),
    Other,
}

fn observe_scalar(
    source: &str,
    execution: &ExecutionContext,
) -> Result<Observed, EvaluationFailure> {
    let expression = parse(source);
    let context = EvaluationContext::with_options(execution, options());
    let result = evaluate_scalar(&expression, &context, &EvaluationLimits::default())?;
    Ok(match result.value() {
        ScalarValue::Number(value) => Observed::Number(*value),
        ScalarValue::Error(error) => Observed::Error(*error),
        _ => Observed::Other,
    })
}

fn observe_value(
    source: &str,
    resolver: &DateResolver,
    execution: &ExecutionContext,
    mode: Mode,
) -> Result<Observed, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0))
        .with_options(options())
        .with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, &Limits::default())?;
    Ok(match result.value() {
        Value::Number(value) => Observed::Number(value),
        Value::Error(error) => Observed::Error(error),
        _ => Observed::Other,
    })
}

fn assert_number(actual: Observed, expected: f64, source: &str) {
    match actual {
        Observed::Number(value) => assert_eq!(value, expected, "{source:?}"),
        other => panic!("{source:?}: expected Number({expected}), got {other:?}"),
    }
}

fn assert_error(actual: Observed, expected: ScalarError, source: &str) {
    assert_eq!(actual, Observed::Error(expected), "{source:?}");
}

fn assert_differential(
    source: &str,
    expected: f64,
    resolver: &DateResolver,
    execution: &ExecutionContext,
) {
    let scalar = observe_scalar(source, execution)
        .unwrap_or_else(|error| panic!("{source:?} scalar failed: {error}"));
    assert_number(scalar, expected, source);
    let value = observe_value(source, resolver, execution, Mode::Scalar)
        .unwrap_or_else(|error| panic!("{source:?} value failed: {error}"));
    assert_number(value, expected, source);
}

#[test]
fn all_twenty_four_functions_have_scalar_value_differential_vectors() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-semantic");
    for (source, expected) in [
        ("=DATE(2020;1;1)", 43_831.0),
        ("=DATEDIF(DATE(2020;1;1);DATE(2021;1;1);\"Y\")", 1.0),
        ("=DATEVALUE(\"2020-02-29\")", 43_890.0),
        ("=DAY(DATE(2020;2;29))", 29.0),
        ("=DAYS(DATE(2020;2;29);DATE(2020;1;1))", 59.0),
        ("=DAYS360(DATE(2020;1;1);DATE(2020;2;1))", 30.0),
        ("=EASTERSUNDAY(2024)", 45_382.0),
        ("=EDATE(DATE(2020;1;31);1)", 43_890.0),
        ("=EOMONTH(DATE(2020;1;15);1)", 43_890.0),
        ("=HOUR(0.5)", 12.0),
        ("=ISOWEEKNUM(DATE(2021;1;1))", 53.0),
        ("=MINUTE(0.5)", 0.0),
        ("=MONTH(DATE(2020;2;29))", 2.0),
        ("=NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;7))", 5.0),
        ("=NOW()", TIMESTAMP_SERIAL),
        ("=SECOND(0.5000115740740741)", 1.0),
        ("=TIME(1;30;0)", 1.0 / 16.0),
        ("=TIMEVALUE(\"12:00:00\")", 0.5),
        ("=TODAY()", TIMESTAMP_SERIAL.floor()),
        ("=WEEKDAY(DATE(2024;1;1);2)", 1.0),
        ("=WEEKNUM(DATE(2024;1;1);21)", 1.0),
        ("=WORKDAY(DATE(2024;1;5);1)", 45_299.0),
        ("=YEAR(DATE(2020;2;29))", 2020.0),
        ("=YEARFRAC(DATE(2020;1;1);DATE(2021;1;1);1)", 1.0),
    ] {
        assert_differential(source, expected, &resolver, &execution);
    }
}

#[test]
fn exact_arity_and_domain_errors_cover_every_date_time_name() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-errors");

    for source in [
        "=DATE(2020;1)",
        "=DATEDIF(1;2)",
        "=DATEVALUE()",
        "=DAY()",
        "=DAYS(1)",
        "=DAYS360(1)",
        "=EASTERSUNDAY(1;2)",
        "=EDATE(1)",
        "=EOMONTH(1)",
        "=HOUR()",
        "=ISOWEEKNUM()",
        "=MINUTE()",
        "=MONTH()",
        "=NETWORKDAYS(1)",
        "=NOW(1)",
        "=SECOND()",
        "=TIME(1;2)",
        "=TIMEVALUE()",
        "=TODAY(1)",
        "=WEEKDAY(1;2;3)",
        "=WEEKNUM(1;2;3)",
        "=WORKDAY(1)",
        "=YEAR()",
        "=YEARFRAC(1)",
    ] {
        assert_error(
            observe_scalar(source, &execution).expect("arity returns a formula value"),
            ScalarError::Value,
            source,
        );
        assert_error(
            observe_value(source, &resolver, &execution, Mode::Scalar)
                .expect("value arity returns a formula value"),
            ScalarError::Value,
            source,
        );
    }

    for (source, expected) in [
        ("=DATE(2020;0;1)", ScalarError::Number),
        ("=DATEDIF(1;2;\"Q\")", ScalarError::Value),
        ("=DATEVALUE(\"not a date\")", ScalarError::Value),
        ("=DAY(\"not a date\")", ScalarError::Value),
        ("=DAYS(\"not a date\";1)", ScalarError::Value),
        ("=DAYS360(1;2;\"bad\")", ScalarError::Value),
        ("=EASTERSUNDAY(1000)", ScalarError::Number),
        ("=EDATE(\"bad\";1)", ScalarError::Value),
        ("=EOMONTH(\"bad\";1)", ScalarError::Value),
        ("=HOUR(\"bad\")", ScalarError::Value),
        ("=ISOWEEKNUM(\"bad\")", ScalarError::Value),
        ("=MINUTE(\"bad\")", ScalarError::Value),
        ("=MONTH(\"bad\")", ScalarError::Value),
        ("=NETWORKDAYS(\"bad\";1)", ScalarError::Value),
        ("=SECOND(\"bad\")", ScalarError::Value),
        ("=TIME(1;\"bad\";0)", ScalarError::Value),
        ("=TIMEVALUE(\"bad\")", ScalarError::Value),
        ("=WEEKDAY(1;4)", ScalarError::Number),
        ("=WEEKNUM(1;3)", ScalarError::Number),
        ("=WEEKNUM(DATE(2024;1;1);1.9)", ScalarError::Number),
        ("=WORKDAY(\"bad\";1)", ScalarError::Value),
        ("=YEAR(\"bad\")", ScalarError::Value),
        ("=YEARFRAC(1;2;5)", ScalarError::Number),
    ] {
        assert_error(
            observe_scalar(source, &execution).expect("domain error is a formula value"),
            expected,
            source,
        );
    }
}

#[test]
fn formula_errors_propagate_through_each_argument_taking_date_function() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-argument-errors");
    for source in [
        "=DATE(NA();1;1)",
        "=DATEDIF(NA();1;\"Y\")",
        "=DATEVALUE(NA())",
        "=DAY(NA())",
        "=DAYS(NA();1)",
        "=DAYS360(NA();1)",
        "=EASTERSUNDAY(NA())",
        "=EDATE(NA();1)",
        "=EOMONTH(NA();1)",
        "=HOUR(NA())",
        "=ISOWEEKNUM(NA())",
        "=MINUTE(NA())",
        "=MONTH(NA())",
        "=NETWORKDAYS(NA();1)",
        "=SECOND(NA())",
        "=TIME(NA();1;1)",
        "=TIMEVALUE(NA())",
        "=WEEKDAY(NA())",
        "=WEEKNUM(NA())",
        "=WORKDAY(NA();1)",
        "=YEAR(NA())",
        "=YEARFRAC(NA();1)",
    ] {
        assert_error(
            observe_scalar(source, &execution).expect("scalar formula error remains a value"),
            ScalarError::NotAvailable,
            source,
        );
        assert_error(
            observe_value(source, &resolver, &execution, Mode::Scalar)
                .expect("value formula error remains a value"),
            ScalarError::NotAvailable,
            source,
        );
    }
}

#[test]
fn every_function_rejects_arguments_outside_its_declared_arity() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-arity-bounds");
    let cases = [
        (Some("=DATE(2020;1)"), "=DATE(2020;1;1;0)"),
        (Some("=DATEDIF(1;2)"), "=DATEDIF(1;2;\"Y\";0)"),
        (Some("=DATEVALUE()"), "=DATEVALUE(\"2020-01-01\";0)"),
        (Some("=DAY()"), "=DAY(1;0)"),
        (Some("=DAYS(1)"), "=DAYS(1;2;0)"),
        (Some("=DAYS360(1)"), "=DAYS360(1;2;FALSE();0)"),
        (None, "=EASTERSUNDAY(2024;1)"),
        (Some("=EDATE(1)"), "=EDATE(1;2;0)"),
        (Some("=EOMONTH(1)"), "=EOMONTH(1;2;0)"),
        (Some("=HOUR()"), "=HOUR(1;0)"),
        (Some("=ISOWEEKNUM()"), "=ISOWEEKNUM(1;0)"),
        (Some("=MINUTE()"), "=MINUTE(1;0)"),
        (Some("=MONTH()"), "=MONTH(1;0)"),
        (Some("=NETWORKDAYS(1)"), "=NETWORKDAYS(1;2;3;4;5)"),
        (None, "=NOW(1)"),
        (Some("=SECOND()"), "=SECOND(1;0)"),
        (Some("=TIME(1;2)"), "=TIME(1;2;3;4)"),
        (Some("=TIMEVALUE()"), "=TIMEVALUE(\"12:00\";0)"),
        (None, "=TODAY(1)"),
        (Some("=WEEKDAY()"), "=WEEKDAY(1;2;3)"),
        (Some("=WEEKNUM()"), "=WEEKNUM(1;2;3)"),
        (Some("=WORKDAY(1)"), "=WORKDAY(1;2;3;4;5)"),
        (Some("=YEAR()"), "=YEAR(1;0)"),
        (Some("=YEARFRAC(1)"), "=YEARFRAC(1;2;0;1)"),
    ];
    for (below, above) in cases {
        for source in below.into_iter().chain(std::iter::once(above)) {
            assert_error(
                observe_scalar(source, &execution).expect("arity refusal is a formula value"),
                ScalarError::Value,
                source,
            );
            assert_error(
                observe_value(source, &resolver, &execution, Mode::Scalar)
                    .expect("value arity refusal is a formula value"),
                ScalarError::Value,
                source,
            );
        }
    }
}

#[test]
fn required_missing_and_optional_empty_slots_keep_distinct_profiles() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-empty-slots");
    for source in [
        "=DATE(;1;1)",
        "=DATEDIF(;1;\"Y\")",
        "=DAYS(;1)",
        "=TIME(1;;0)",
        "=WORKDAY(;1)",
        "=YEARFRAC(;1)",
    ] {
        assert_error(
            observe_scalar(source, &execution).expect("required missing slot is a value"),
            ScalarError::Value,
            source,
        );
        assert_error(
            observe_value(source, &resolver, &execution, Mode::Scalar)
                .expect("value required missing slot is a value"),
            ScalarError::Value,
            source,
        );
    }
    for (source, expected) in [
        ("=DAYS360(DATE(2024;1;1);DATE(2024;2;1);)", 30.0),
        ("=NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;7);;)", 5.0),
        ("=WORKDAY(DATE(2024;1;5);1;;)", 45_299.0),
        ("=WEEKDAY(DATE(2024;1;1);)", 2.0),
        ("=WEEKNUM(DATE(2024;1;1);)", 1.0),
        ("=YEARFRAC(DATE(2020;1;1);DATE(2021;1;1);)", 1.0),
    ] {
        assert_differential(source, expected, &resolver, &execution);
    }
}

#[test]
fn datedif_formats_trim_case_and_apply_signed_month_remainders() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-datedif-profile");
    for (source, expected) in [
        ("=DATEDIF(43524;43890;\" y \")", 1.0),
        ("=DATEDIF(43831;43890;\"m\")", 1.0),
        ("=DATEDIF(43831;43890;\" D \")", 59.0),
        ("=DATEDIF(43831;43890;\"md\")", 28.0),
        ("=DATEDIF(43861;43889;\"YM\")", 0.0),
        ("=DATEDIF(43861;44197;\"YM\")", 11.0),
    ] {
        assert_differential(source, expected, &resolver, &execution);
    }
    let source = "=DATEDIF(43831;43890;\" \" )";
    assert_error(
        observe_scalar(source, &execution).expect("empty format is a formula value"),
        ScalarError::Value,
        source,
    );
}

#[test]
fn time_parameters_accept_finite_values_outside_the_date_domain() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-time-domain");
    for (source, expected) in [
        ("=HOUR(4000000.5)", 12.0),
        ("=TIMEVALUE(\"4000000.5\")", 4_000_000.5),
        ("=HOUR(-0.5)", 12.0),
        ("=HOUR(-5e-324)", 23.0),
        ("=MINUTE(-1/86400)", 59.0),
        ("=SECOND(-1/86400)", 59.0),
        ("=MINUTE(-1/256)", 54.0),
        ("=SECOND(-1/256)", 22.0),
        ("=TIME(-0.000001;0;0)", -0.000001 / 24.0),
    ] {
        assert_differential(source, expected, &resolver, &execution);
    }
}

#[test]
fn date_bounds_fractional_floor_and_zero_offsets_follow_the_profile() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-bounds");
    for (source, expected) in [
        ("=DATE(1;1;1)", -693_593.0),
        ("=DATE(9999;12;31)", 2_958_465.0),
        ("=DATE(2020.9;1.9;1.9)", 43_831.0),
        ("=DATEVALUE(\"0001-01-01\")", -693_593.0),
        ("=DATEVALUE(\"9999-12-31\")", 2_958_465.0),
        ("=DAY(43890.75)", 29.0),
        ("=MONTH(43890.75)", 2.0),
        ("=YEAR(43890.75)", 2020.0),
        ("=DAYS360(43890.75;43891.25)", 1.0),
        ("=EDATE(43890.75;0)", 43_890.0),
        ("=EOMONTH(43890.75;0)", 43_890.0),
        ("=WORKDAY(45298.75;0)", 45_298.75),
    ] {
        assert_differential(source, expected, &resolver, &execution);
    }
    for source in ["=DAY(2958466)", "=YEAR(-693594)"] {
        assert_error(
            observe_scalar(source, &execution).expect("date bound failure is a formula value"),
            ScalarError::Number,
            source,
        );
    }
}

#[test]
fn retained_holiday_formula_error_supersedes_generated_scalar_conversion_error() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-error-precedence");
    let source = "=NETWORKDAYS(\"bad\";45292;{NA()})";
    assert_error(
        observe_value(source, &resolver, &execution, Mode::Scalar)
            .expect("formula errors remain formula values"),
        ScalarError::NotAvailable,
        source,
    );
}

#[test]
fn fixed_parsers_accept_grouped_numeric_fallback_and_mixed_fractions() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-parsing");
    for (source, expected) in [
        ("=DATEVALUE(\"1,234.5\")", 1_234.0),
        ("=TIMEVALUE(\"1,234.5\")", 1_234.5),
        ("=DATEVALUE(\"2006-05-21 12:34:56.5\")", 38_858.0),
        (
            "=TIMEVALUE(\"2006-05-21 12:34:56.5\")",
            12.0 / 24.0 + 34.0 / 1440.0 + 56.5 / 86_400.0,
        ),
    ] {
        assert_differential(source, expected, &resolver, &execution);
    }

    let mut grouped = String::from("0");
    for _ in 0..200 {
        grouped.push_str(",000");
    }
    grouped.push_str(",001.5");
    let numeric = grouped
        .replace(',', "")
        .parse::<f64>()
        .expect("finite grouped number");
    let source = format!("=TIMEVALUE(\"{grouped}\")");
    assert_differential(&source, numeric, &resolver, &execution);
}

#[test]
fn numeric_fallback_covers_exponent_percent_currency_and_fraction_forms() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-value-fallback");
    for (source, expected) in [
        ("=DATEVALUE(\"1.25E2\")", 125.0),
        ("=DATEVALUE(\"50%\")", 0.0),
        ("=DATEVALUE(\"$123.5\")", 123.0),
        ("=DATEVALUE(\"1 1/2\")", 1.0),
        ("=TIMEVALUE(\"1.25E2\")", 125.0),
        ("=TIMEVALUE(\"50%\")", 0.5),
        ("=TIMEVALUE(\"$123.5\")", 123.5),
        ("=TIMEVALUE(\"(\u{0024}123.5)\")", -123.5),
        ("=TIMEVALUE(\"1 1/2\")", 1.5),
    ] {
        assert_differential(source, expected, &resolver, &execution);
    }

    for source in ["=DATEVALUE(\"1,23\")", "=TIMEVALUE(\"1,23\")"] {
        assert_error(
            observe_scalar(source, &execution).expect("malformed numeric fallback is a value"),
            ScalarError::Value,
            source,
        );
        assert_error(
            observe_value(source, &resolver, &execution, Mode::Scalar)
                .expect("value malformed numeric fallback is a value"),
            ScalarError::Value,
            source,
        );
    }

    for source in [
        "=DATEVALUE(\"1e309\")",
        "=DATEVALUE(\"-693594\")",
        "=DATEVALUE(\"2958466\")",
        "=TIMEVALUE(\"1e309\")",
    ] {
        assert_error(
            observe_scalar(source, &execution).expect("numeric fallback domain failure is a value"),
            ScalarError::Number,
            source,
        );
        assert_error(
            observe_value(source, &resolver, &execution, Mode::Scalar)
                .expect("value numeric fallback domain failure is a value"),
            ScalarError::Number,
            source,
        );
    }
}

#[test]
fn yearfrac_and_iso_week_edges_follow_profile_procedures() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-calendar-edges");
    for (source, expected) in [
        ("=YEARFRAC(DATE(2020;2;29);DATE(2021;2;28);0)", 1.0),
        ("=YEARFRAC(DATE(2020;2;28);DATE(2020;3;1);1)", 2.0 / 366.0),
        (
            "=YEARFRAC(DATE(2019;1;1);DATE(2021;1;1);1)",
            2.0009124087591244,
        ),
        ("=ISOWEEKNUM(DATE(2021;1;1))", 53.0),
        ("=ISOWEEKNUM(DATE(2021;1;2))", 53.0),
        ("=ISOWEEKNUM(DATE(2021;1;3))", 53.0),
        ("=ISOWEEKNUM(DATE(2021;1;4))", 1.0),
    ] {
        assert_differential(source, expected, &resolver, &execution);
    }
}

#[test]
fn yearfrac_defaults_floors_dates_and_handles_thirty_first_rules() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-yearfrac-edges");
    for (source, expected) in [
        ("=YEARFRAC(43831;43890)", 58.0 / 360.0),
        ("=YEARFRAC(43831;43890;)", 58.0 / 360.0),
        ("=YEARFRAC(43831.75;43832.25;3)", 1.0 / 365.0),
        ("=YEARFRAC(44227;44255;4)", 28.0 / 360.0),
    ] {
        assert_differential(source, expected, &resolver, &execution);
    }
}

#[test]
fn timestamp_api_is_explicit_and_volatile_functions_are_deterministic() {
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-timestamp");
    let timestamp = CalculationTimestamp::from_serial(TIMESTAMP_SERIAL).expect("valid timestamp");
    assert_eq!(timestamp.serial(), TIMESTAMP_SERIAL);
    assert_eq!(timestamp, timestamp);
    for serial in [-f64::MIN_POSITIVE, -f64::from_bits(1)] {
        let roundtrip = CalculationTimestamp::from_serial(serial)
            .expect("finite near-zero serial remains in the date profile");
        assert_eq!(roundtrip.serial().to_bits(), serial.to_bits());
    }
    assert!(CalculationTimestamp::from_ymd_hms(9999, 12, 31, 23, 59, 59.999999).is_err());
    assert!(CalculationTimestamp::from_serial(f64::NAN).is_err());
    assert!(CalculationTimestamp::from_serial(2_958_466.0).is_err());

    let expression = parse("=NOW()");
    let context = EvaluationContext::with_options(
        &execution,
        EvaluationOptions::default().with_calculation_timestamp(timestamp),
    );
    let result = evaluate_scalar(&expression, &context, &EvaluationLimits::default())
        .expect("timestamp-backed NOW");
    assert_eq!(result.value(), &ScalarValue::Number(TIMESTAMP_SERIAL));

    for source in ["=NOW()", "=TODAY()", "=EASTERSUNDAY()"] {
        let expression = parse(source);
        let result = evaluate_scalar(
            &expression,
            &EvaluationContext::new(&execution),
            &EvaluationLimits::default(),
        );
        assert!(
            matches!(
                result,
                Err(EvaluationFailure::Unsupported(
                    UnsupportedKind::CalculationClock
                ))
            ),
            "{source} must refuse without a timestamp"
        );
    }
}

#[test]
fn matrix_date_arguments_preserve_shape_and_elementwise_results() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-matrix");

    for (source, expected) in [
        ("=DATE({2020|2021};1;1)", vec![43_831.0, 44_197.0]),
        ("=TIME({1|2};0;0)", vec![1.0 / 24.0, 1.0 / 12.0]),
        (
            "=WEEKDAY({DATE(2024;1;1)|DATE(2024;1;7)};2)",
            vec![1.0, 7.0],
        ),
        (
            "=IF({TRUE()|FALSE()};DATE(2020;1;1);DATE(2021;1;1))",
            vec![43_831.0, 44_197.0],
        ),
    ] {
        let expression = parse(source);
        let context = Context::new(&execution, Position::new("Main", 0, 0))
            .with_options(options())
            .with_mode(Mode::Matrix);
        let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
            .unwrap_or_else(|error| panic!("{source:?}: {error}"));
        let array = result.as_array().expect("matrix date result");
        assert_eq!(array.shape().rows(), 2, "{source:?} rows");
        assert_eq!(array.shape().columns(), 1, "{source:?} columns");
        for (index, expected) in expected.into_iter().enumerate() {
            assert_eq!(
                array.get(index),
                Some(Value::Number(expected)),
                "{source:?}[{index}]"
            );
        }
    }
}

#[test]
fn projected_datevalue_reference_preserves_each_date_and_reads_once_per_cell() {
    let mut resolver = DateResolver::standard();
    resolver.set(0, 0, FixtureCell::Text("2020-01-01".to_owned()));
    resolver.set(1, 0, FixtureCell::Text("2020-01-02".to_owned()));
    let (_budget, _cancellation, execution) =
        execution("ods-formula-date-time-projected-datevalue");
    let expression = parse("=IF({TRUE()|TRUE()};DATEVALUE([.A1:.A2]);0)");
    let context = Context::new(&execution, Position::new("Main", 0, 0))
        .with_options(options())
        .with_mode(Mode::Matrix);
    let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
        .expect("projected DATEVALUE reference");
    let array = result.as_array().expect("projected DATEVALUE array");
    assert_eq!(array.shape().rows(), 2);
    assert_eq!(array.shape().columns(), 1);
    assert_eq!(array.get(0), Some(Value::Number(43_831.0)));
    assert_eq!(array.get(1), Some(Value::Number(43_832.0)));
    assert_eq!(resolver.reads(), 2);
}

#[test]
fn complete_holiday_and_workweek_sequences_survive_projected_if() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-sequences");
    let expression = parse(
        "=IF({TRUE()|TRUE()};NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;12);[.A1:.A2];[.B1:.B7]);0)",
    );
    let context = Context::new(&execution, Position::new("Main", 0, 0))
        .with_options(options())
        .with_mode(Mode::Matrix);
    let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
        .expect("projected sequence reducer");
    let array = result.as_array().expect("projected sequence array");
    assert_eq!(array.shape().rows(), 2);
    assert_eq!(array.shape().columns(), 1);
    assert_eq!(array.get(0), Some(Value::Number(8.0)));
    assert_eq!(array.get(1), Some(Value::Number(8.0)));
    assert!(
        resolver.reads() >= 9,
        "both complete sequences must be consumed"
    );
    let order = resolver.read_order();
    for coordinate in [(0, 0), (1, 0), (0, 1), (6, 1)] {
        assert!(
            order
                .iter()
                .any(|(_, row, column)| (*row, *column) == coordinate),
            "missing sequence coordinate {coordinate:?}"
        );
    }
}

#[test]
fn nested_munit_date_selector_remains_position_sensitive() {
    let resolver = DateResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-munit");
    let expression = parse("=IF({TRUE()|TRUE()};WEEKDAY(DATE(2024;1;1);SUM(MUNIT([.G1:.G2])));0)");
    let context = Context::new(&execution, Position::new("Main", 0, 0))
        .with_options(options())
        .with_mode(Mode::Matrix);
    let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
        .expect("position-sensitive date selector");
    let array = result.as_array().expect("position-sensitive date array");
    assert_eq!(array.shape().rows(), 2);
    assert_eq!(array.shape().columns(), 1);
    assert_eq!(array.get(0), Some(Value::Number(2.0)));
    assert_eq!(array.get(1), Some(Value::Number(1.0)));
}

#[test]
fn value_profile_does_not_widen_existing_value_conversion() {
    let (_budget, _cancellation, execution) = execution("ods-formula-date-time-value-compat");
    for (source, expected) in [
        ("=VALUE(\"2006-05-21\")", 38_858.0),
        ("=VALUE(\"2006-05-21 12:00\")", 38_858.5),
        ("=VALUE(\"2:00\")", 1.0 / 12.0),
    ] {
        assert_number(
            observe_scalar(source, &execution).expect("existing VALUE result"),
            expected,
            source,
        );
    }
    assert_error(
        observe_scalar("=VALUE(\"1899-12-29\")", &execution).expect("VALUE domain result"),
        ScalarError::Number,
        "=VALUE(\"1899-12-29\")",
    );
}
