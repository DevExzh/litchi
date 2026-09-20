//! Resource and failure-boundary coverage for the ODF 1.4 date/time family.
//! Resolver-backed sequence arguments are intentionally observable so tests
//! can distinguish complete consumption from projected scalar intersection.

use std::{
    cell::{Cell, RefCell},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource, SourceVersion,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Limits, Mode, Position, Resolver, SheetExtent, Value,
    },
    evaluation::{
        CalculationTimestamp, EvaluationFailure, EvaluationOptions, ScalarError, UnsupportedKind,
    },
    expression::Expression,
};

#[derive(Debug, Clone)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
    Unsupported,
}

#[derive(Debug)]
struct LimitsResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(usize, usize)>>,
    cancel_after_read: Option<CancellationSource>,
    cancel_on_read: Option<(usize, CancellationSource)>,
    fail_after_read: Cell<Option<(usize, UnsupportedKind)>>,
    source_versions: Cell<Option<(SourceVersion, SourceVersion)>>,
    source_version_calls: Cell<usize>,
}

impl LimitsResolver {
    fn standard() -> Self {
        let rows = 32;
        let columns = 16;
        let mut resolver = Self {
            rows,
            columns,
            cells: vec![FixtureCell::Empty; rows * columns],
            reads: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
            cancel_after_read: None,
            cancel_on_read: None,
            fail_after_read: Cell::new(None),
            source_versions: Cell::new(None),
            source_version_calls: Cell::new(0),
        };
        for (row, value) in [45_292.0, 45_299.0, 45_300.0, 45_301.0]
            .into_iter()
            .enumerate()
        {
            resolver.set(row, 0, FixtureCell::Number(value));
        }
        for (row, value) in [true, false, false, false, false, false, true]
            .into_iter()
            .enumerate()
        {
            resolver.set(row, 1, FixtureCell::Logical(value));
        }
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

    fn clear(&self) {
        self.reads.set(0);
        self.read_order.borrow_mut().clear();
        self.fail_after_read.set(None);
        self.source_version_calls.set(0);
    }

    fn cancel_after_read(&mut self, cancellation: &CancellationSource) {
        self.cancel_after_read = Some(cancellation.clone());
    }

    fn cancel_on_read(&mut self, read: usize, cancellation: &CancellationSource) {
        self.cancel_on_read = Some((read, cancellation.clone()));
    }

    fn fail_after_read(&self, successful_reads: usize, kind: UnsupportedKind) {
        self.fail_after_read.set(Some((successful_reads, kind)));
    }

    fn set_source_versions(&self, expected: SourceVersion, observed: SourceVersion) {
        self.source_versions.set(Some((expected, observed)));
        self.source_version_calls.set(0);
    }
}

impl Resolver for LimitsResolver {
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
        self.read_order.borrow_mut().push((row, column));
        if let Some((after, kind)) = self.fail_after_read.get() {
            if self.reads.get() > after {
                return Err(EvaluationFailure::Unsupported(kind));
            }
        }
        if let Some(cancellation) = &self.cancel_after_read {
            cancellation.cancel();
        }
        if let Some((read, cancellation)) = &self.cancel_on_read {
            if self.reads.get() >= *read {
                cancellation.cancel();
            }
        }
        if sheet != "Main" || row >= self.rows || column >= self.columns {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        Ok(match &self.cells[row * self.columns + column] {
            FixtureCell::Empty => CellRead::Empty,
            FixtureCell::Number(value) => CellRead::Number(*value),
            FixtureCell::Logical(value) => CellRead::Logical(*value),
            FixtureCell::Text(value) => CellRead::Text(value.as_str()),
            FixtureCell::Error(error) => CellRead::Error(*error),
            FixtureCell::Unsupported => CellRead::Unsupported,
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

    fn source_version(
        &self,
        _execution: &ExecutionContext,
    ) -> Result<Option<SourceVersion>, EvaluationFailure> {
        let Some((expected, observed)) = self.source_versions.get() else {
            return Ok(None);
        };
        let call = self.source_version_calls.get();
        self.source_version_calls.set(call.saturating_add(1));
        Ok(Some(if call == 0 { expected } else { observed }))
    }
}

fn make_execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
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
enum Observed {
    Number(f64),
    Error(ScalarError),
    Other,
}

fn evaluate_source(
    source: &str,
    resolver: &LimitsResolver,
    execution: &ExecutionContext,
    limits: &Limits,
) -> Result<Observed, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(match result.value() {
        Value::Number(value) => Observed::Number(value),
        Value::Error(error) => Observed::Error(error),
        Value::Array(array) if array.shape().rows() == 1 && array.shape().columns() == 1 => {
            match array.get(0) {
                Some(Value::Number(value)) => Observed::Number(value),
                Some(Value::Error(error)) => Observed::Error(error),
                _ => Observed::Other,
            }
        },
        _ => Observed::Other,
    })
}

fn evaluate_source_with_options(
    source: &str,
    resolver: &LimitsResolver,
    execution: &ExecutionContext,
    limits: &Limits,
    options: EvaluationOptions,
) -> Result<Observed, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0))
        .with_options(options)
        .with_mode(Mode::Matrix);
    let result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(match result.value() {
        Value::Number(value) => Observed::Number(value),
        Value::Error(error) => Observed::Error(error),
        Value::Array(array) if array.shape().rows() == 1 && array.shape().columns() == 1 => {
            match array.get(0) {
                Some(Value::Number(value)) => Observed::Number(value),
                Some(Value::Error(error)) => Observed::Error(error),
                _ => Observed::Other,
            }
        },
        _ => Observed::Other,
    })
}

#[test]
fn reference_cell_and_array_limits_fail_before_reads_and_release_memory() {
    let resolver = LimitsResolver::standard();
    let (budget, _cancellation, execution) = make_execution("ods-formula-date-time-cell-limit");
    let error = evaluate_source(
        "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;12);[.A1:.A2])",
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(1),
    )
    .expect_err("reference-cell admission must be bounded");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let resolver = LimitsResolver::standard();
    let (array_budget, _cancellation, array_execution) =
        make_execution("ods-formula-date-time-array-limit");
    let error = evaluate_source(
        "=DATE({2020|2021|2022};1;1)",
        &resolver,
        &array_execution,
        &Limits::default().with_max_array_cells(2),
    )
    .expect_err("matrix output admission must be bounded");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(array_budget.used(Resource::Memory), 0);
}

#[test]
fn reference_list_sequence_refusal_is_read_free_but_invalid_reference_length_is_eager() {
    let resolver = LimitsResolver::standard();
    let (_budget, _cancellation, execution) = make_execution("ods-formula-date-time-shape");
    let result = evaluate_source(
        "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;12);([.A1]~[.A2]))",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("list refusal is a formula value");
    assert_eq!(result, Observed::Error(ScalarError::Value));
    assert_eq!(resolver.reads(), 0);

    resolver.clear();
    let result = evaluate_source(
        "=NETWORKDAYS(NA();1;([.A1]~[.A2]))",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("earlier formula error remains a formula value");
    assert_eq!(result, Observed::Error(ScalarError::NotAvailable));
    assert_eq!(resolver.reads(), 0);

    resolver.clear();
    let result = evaluate_source(
        "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;12);[.A1:.A2];[.B1:.B2])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("referenced workweek length is a formula value");
    assert_eq!(result, Observed::Error(ScalarError::Value));
    assert_eq!(
        resolver.reads(),
        4,
        "both referenced sequences are consumed"
    );
}

#[test]
fn formula_errors_are_retained_while_later_sequence_failures_supersede_them() {
    let mut resolver = LimitsResolver::standard();
    resolver.set(0, 0, FixtureCell::Error(ScalarError::NotAvailable));
    resolver.fail_after_read(1, UnsupportedKind::CellValue);
    let (_budget, _cancellation, execution) = make_execution("ods-formula-date-time-error-order");
    let error = evaluate_source(
        "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;12);[.A1:.A4])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("typed sequence failure supersedes a retained formula error");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), 2);

    resolver.clear();
    let result = evaluate_source(
        "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;12);[.A1:.A4])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("formula error remains a formula value");
    assert_eq!(result, Observed::Error(ScalarError::NotAvailable));
    assert_eq!(resolver.reads(), 4);
}

#[test]
fn cancellation_and_source_fences_publish_no_partial_date_result() {
    let (budget, cancellation, execution) = make_execution("ods-formula-date-time-cancel");
    let mut resolver = LimitsResolver::standard();
    resolver.cancel_after_read(&cancellation);
    let error = evaluate_source(
        "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;12);[.A1:.A4])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("cancellation must fence date sequence publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);

    let resolver = LimitsResolver::standard();
    let expected = SourceVersion::new(0x4441_5445, 0);
    let observed = SourceVersion::new(0x4441_5445, 1);
    resolver.set_source_versions(expected, observed);
    let (_budget, _cancellation, source_execution) = make_execution("ods-formula-date-time-source");
    let error = evaluate_source(
        "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;12);[.A1:.A4])",
        &resolver,
        &source_execution,
        &Limits::default(),
    )
    .expect_err("source change must fence date result publication");
    assert!(matches!(
        error,
        EvaluationFailure::SourceChanged {
            expected: got_expected,
            observed: got_observed
        } if got_expected == expected && got_observed == observed
    ));
    assert!(resolver.reads() > 0);
}

#[test]
fn borrowed_date_text_respects_text_budget_and_typed_failures_are_not_catchable() {
    let mut resolver = LimitsResolver::standard();
    resolver.set(0, 0, FixtureCell::Text("2020-02-29".to_owned()));
    let (_budget, _cancellation, execution) = make_execution("ods-formula-date-time-text");
    let result = evaluate_source(
        "=DATEVALUE([.A1])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("borrowed date text");
    assert_eq!(result, Observed::Number(43_890.0));
    assert_eq!(resolver.reads(), 1);

    resolver.clear();
    let error = evaluate_source(
        "=DATEVALUE([.A1])",
        &resolver,
        &execution,
        &Limits::default().with_max_text_bytes(3),
    )
    .expect_err("date text budget must remain typed");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "unexpected text budget failure: {error:?}"
    );

    resolver.clear();
    resolver.set(0, 0, FixtureCell::Unsupported);
    resolver.fail_after_read(0, UnsupportedKind::CellValue);
    let error = evaluate_source(
        "=IFERROR(DATEVALUE([.A1]);0)",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("IFERROR must not catch a provider failure");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
}

#[test]
fn long_date_intervals_are_bounded_by_work_even_without_reference_reads() {
    let resolver = LimitsResolver::standard();
    let (budget, _cancellation, execution) = make_execution("ods-formula-date-time-work");
    let error = evaluate_source(
        "=NETWORKDAYS(DATE(1900;1;1);DATE(9999;12;31))",
        &resolver,
        &execution,
        &Limits::default().with_max_steps(16),
    )
    .expect_err("long date iteration must charge Work");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work
    ));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn workday_consumes_holiday_and_workweek_sequences_before_zero_and_all_off_results() {
    let mut resolver = LimitsResolver::standard();
    for row in 0..16 {
        resolver.set(row, 0, FixtureCell::Number(40_000.0 + row as f64));
    }
    for row in 0..7 {
        resolver.set(row, 1, FixtureCell::Logical(true));
    }

    let (budget, _cancellation, execution) =
        make_execution("ods-formula-date-time-workday-sequence-complete");
    let source = "=WORKDAY(DATE(2024;1;5);0;[.A1:.A16];[.B1:.B7])";
    let result = evaluate_source(source, &resolver, &execution, &Limits::default())
        .expect("zero offset remains a value after complete sequence scans");
    assert_eq!(result, Observed::Number(45_296.0));
    assert_eq!(
        resolver.reads(),
        23,
        "holiday and workweek ranges are eager"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    resolver.clear();
    let result = evaluate_source(
        "=WORKDAY(DATE(2024;1;5);1;[.A1:.A16];[.B1:.B7])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("all-off workweek returns a formula result");
    assert_eq!(result, Observed::Error(ScalarError::Number));
    assert_eq!(
        resolver.reads(),
        23,
        "all-off still consumes both sequences"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    resolver.clear();
    resolver.set(0, 0, FixtureCell::Error(ScalarError::NotAvailable));
    resolver.fail_after_read(16, UnsupportedKind::CellValue);
    let error = evaluate_source(source, &resolver, &execution, &Limits::default())
        .expect_err("a later typed sequence failure supersedes a retained formula error");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(
        resolver.reads(),
        17,
        "the failure occurs after all holidays"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn matrix_workday_typed_failure_publishes_no_partial_result() {
    let mut resolver = LimitsResolver::standard();
    resolver.fail_after_read(0, UnsupportedKind::CellValue);
    let (budget, _cancellation, execution) = make_execution("ods-formula-date-time-matrix-failure");
    let error = evaluate_source(
        "=IF({TRUE()|TRUE()};WORKDAY(DATE(2024;1;5);0;[.A1:.A4];[.B1:.B7]);0)",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("typed failure must abort matrix publication");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);

    resolver.clear();
    resolver.set(0, 0, FixtureCell::Text("2020-01-01".to_owned()));
    resolver.set(1, 0, FixtureCell::Text("2020-01-02".to_owned()));
    resolver.fail_after_read(1, UnsupportedKind::CellValue);
    let error = evaluate_source(
        "=IF({TRUE()|TRUE()};DATEVALUE([.A1:.A2]);0)",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("a later matrix cell failure must discard the earlier cell result");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), 2);
    assert_eq!(resolver.read_order.borrow().as_slice(), &[(0, 0), (1, 0)]);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn long_grouped_parser_matrix_and_holiday_storage_release_resources() {
    let mut grouped = "0".repeat(1_100);
    grouped.push_str("45,292");

    let mut resolver = LimitsResolver::standard();
    resolver.set(0, 0, FixtureCell::Text(grouped.clone()));
    let (budget, _cancellation, execution) = make_execution("ods-formula-date-time-long-parser");
    let result = evaluate_source(
        "=DATEVALUE([.A1])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("long grouped date text is parsed with bounded scratch");
    assert_eq!(result, Observed::Number(45_292.0));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);

    resolver.clear();
    resolver.set(1, 0, FixtureCell::Text(grouped.clone()));
    {
        let expression = parse("=DATEVALUE([.A1:.A2])");
        let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
        let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
            .expect("matrix parsing of long grouped text");
        let array = result.as_array().expect("matrix DATEVALUE result");
        assert_eq!(array.shape().rows(), 2);
        assert_eq!(array.shape().columns(), 1);
        assert_eq!(array.get(0), Some(Value::Number(45_292.0)));
        assert_eq!(array.get(1), Some(Value::Number(45_292.0)));
    }
    assert_eq!(resolver.reads(), 2);
    assert_eq!(budget.used(Resource::Memory), 0);

    resolver.clear();
    resolver.set(0, 0, FixtureCell::Text(grouped));
    let error = evaluate_source(
        "=DATEVALUE([.A1])",
        &resolver,
        &execution,
        &Limits::default().with_max_storage_bytes(1_105),
    )
    .expect_err("scratch reservation must be bounded before parsing");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory
    ));
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn explicit_timestamp_contexts_change_results_and_remain_fenced() {
    let timestamp_one = CalculationTimestamp::from_serial(46_000.5).expect("timestamp one");
    let timestamp_two = CalculationTimestamp::from_serial(46_001.25).expect("timestamp two");
    let options_one = EvaluationOptions::default().with_calculation_timestamp(timestamp_one);
    let options_two = EvaluationOptions::default().with_calculation_timestamp(timestamp_two);

    let resolver = LimitsResolver::standard();
    let (budget, _cancellation, execution) =
        make_execution("ods-formula-date-time-timestamp-contexts");
    let first = evaluate_source_with_options(
        "=NOW()",
        &resolver,
        &execution,
        &Limits::default(),
        options_one,
    )
    .expect("first explicit timestamp");
    let second = evaluate_source_with_options(
        "=NOW()",
        &resolver,
        &execution,
        &Limits::default(),
        options_two,
    )
    .expect("second explicit timestamp");
    assert_eq!(first, Observed::Number(46_000.5));
    assert_eq!(second, Observed::Number(46_001.25));
    assert_ne!(
        first, second,
        "timestamp identity must affect volatile values"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let resolver = LimitsResolver::standard();
    let expected = SourceVersion::new(0x5449_4d45, 0);
    let observed = SourceVersion::new(0x5449_4d45, 1);
    resolver.set_source_versions(expected, observed);
    let (source_budget, _cancellation, source_execution) =
        make_execution("ods-formula-date-time-timestamp-source");
    let error = evaluate_source_with_options(
        "=TODAY()",
        &resolver,
        &source_execution,
        &Limits::default(),
        options_two,
    )
    .expect_err("source change must fence timestamp-backed publication");
    assert!(matches!(
        error,
        EvaluationFailure::SourceChanged {
            expected: got_expected,
            observed: got_observed
        } if got_expected == expected && got_observed == observed
    ));
    assert_eq!(source_budget.used(Resource::Memory), 0);

    let resolver = LimitsResolver::standard();
    let (cancel_budget, cancellation, cancel_execution) =
        make_execution("ods-formula-date-time-timestamp-cancel");
    cancellation.cancel();
    let error = evaluate_source_with_options(
        "=NOW()",
        &resolver,
        &cancel_execution,
        &Limits::default(),
        options_one,
    )
    .expect_err("final cancellation must prevent timestamp publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(cancel_budget.used(Resource::Memory), 0);
}

#[test]
fn workday_cancellation_follows_complete_sequences_and_work_is_bounded() {
    let mut resolver = LimitsResolver::standard();
    for row in 0..16 {
        resolver.set(row, 0, FixtureCell::Number(40_000.0 + row as f64));
    }
    let (budget, cancellation, execution) = make_execution("ods-formula-date-time-workday-cancel");
    resolver.cancel_on_read(23, &cancellation);
    let error = evaluate_source(
        "=WORKDAY(DATE(2024;1;5);1;[.A1:.A16];[.B1:.B7])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("WORKDAY must check cancellation after complete sequence consumption");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 23);
    let expected_order: Vec<_> = (0..16)
        .map(|row| (row, 0))
        .chain((0..7).map(|row| (row, 1)))
        .collect();
    assert_eq!(*resolver.read_order.borrow(), expected_order);
    assert_eq!(budget.used(Resource::Memory), 0);

    let resolver = LimitsResolver::standard();
    let (budget, _cancellation, execution) = make_execution("ods-formula-date-time-workday-work");
    let error = evaluate_source(
        "=WORKDAY(DATE(2024;1;5);1024)",
        &resolver,
        &execution,
        &Limits::default().with_max_steps(16),
    )
    .expect_err("WORKDAY date stepping must charge bounded work");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work
    ));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}
