//! Resource, cancellation, source-freshness, typed-failure, and cache-boundary
//! coverage for the paired statistics and simple regression functions.
//!
//! The resolver makes every cell read observable.  Shape and pseudotype
//! refusals therefore prove that geometry is checked before provider access,
//! while the larger cases prove cumulative two-array limits and invariant
//! FORECAST-fit reuse.

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
        self, CellRead, Context, Evaluated, Limits, Mode, Position, Resolver, SheetExtent, Value,
    },
    evaluation::{EvaluationFailure, ScalarError, UnsupportedKind},
    expression::Expression,
};

const X: &str = "[.A1:.A4]";
const Y: &str = "[.B1:.B4]";

#[derive(Debug, Clone)]
enum FixtureCell {
    Empty,
    Number(f64),
    Text(String),
    Error(ScalarError),
}

#[derive(Debug)]
struct LimitsResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_coordinates: RefCell<Vec<(usize, usize)>>,
    cancel_after_read: Option<CancellationSource>,
    failure_at_read: Option<usize>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl LimitsResolver {
    fn standard(rows: usize) -> Self {
        let columns = 8;
        let mut resolver = Self {
            rows,
            columns,
            cells: (0..rows * columns).map(|_| FixtureCell::Empty).collect(),
            reads: Cell::new(0),
            read_coordinates: RefCell::new(Vec::new()),
            cancel_after_read: None,
            failure_at_read: None,
            source_versions: None,
            source_version_calls: Cell::new(0),
        };
        for row in 0..rows {
            let x = (row + 1) as f64;
            resolver.set(row, 0, FixtureCell::Number(x));
            resolver.set(row, 1, FixtureCell::Number(2.0 * x + 1.0));
        }
        if rows >= 2 {
            resolver.set(0, 6, FixtureCell::Number(5.0));
            resolver.set(1, 6, FixtureCell::Number(6.0));
        }
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

    fn with_cancel_after_read(mut self, cancellation: &CancellationSource) -> Self {
        self.cancel_after_read = Some(cancellation.clone());
        self
    }

    fn with_failure_at_read(mut self, read: usize) -> Self {
        self.failure_at_read = Some(read);
        self
    }

    fn set_source_versions(&mut self, expected: SourceVersion, observed: SourceVersion) {
        self.source_versions = Some((expected, observed));
        self.source_version_calls.set(0);
    }
}

impl Resolver for LimitsResolver {
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
        let read = self.reads.get().saturating_add(1);
        self.reads.set(read);
        self.read_coordinates.borrow_mut().push((row, column));
        if let Some(cancellation) = &self.cancel_after_read {
            cancellation.cancel();
        }
        if self.failure_at_read == Some(read) {
            return Err(EvaluationFailure::Unsupported(UnsupportedKind::CellValue));
        }
        if !matches!(sheet, "Main" | "Data" | "Archive")
            || row >= self.rows
            || column >= self.columns
        {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        // Additional planes are exposed for the known 3-D refusal, but the
        // paired profile rejects that descriptor before this branch is read.
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
enum LimitResult {
    Number(f64),
    Error(ScalarError),
    Other,
}

fn evaluate_source(
    source: &str,
    resolver: &LimitsResolver,
    execution: &ExecutionContext,
    limits: &Limits,
) -> Result<LimitResult, EvaluationFailure> {
    evaluate_source_mode(source, resolver, execution, Mode::Matrix, limits)
}

fn evaluate_source_mode(
    source: &str,
    resolver: &LimitsResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<LimitResult, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(match result.value() {
        Value::Number(value) => LimitResult::Number(value),
        Value::Error(error) => LimitResult::Error(error),
        _ => LimitResult::Other,
    })
}

fn evaluate<'a>(
    expression: &'a Expression,
    resolver: &'a LimitsResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    value::evaluate(expression, resolver, &context, limits)
}

fn assert_number(result: LimitResult, expected: f64, source: &str) {
    let LimitResult::Number(actual) = result else {
        panic!("{source:?}: expected Number({expected}), got {result:?}");
    };
    let tolerance = expected.abs().max(actual.abs()).max(1.0) * 64.0 * f64::EPSILON;
    assert!(
        (actual - expected).abs() <= tolerance,
        "{source:?}: {actual} != {expected} (tolerance {tolerance})"
    );
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
            Some(Value::Number(actual)) => assert_number(
                LimitResult::Number(actual),
                expected,
                &format!("{source:?}[{index}]"),
            ),
            other => panic!("{source:?}[{index}]: expected Number, got {other:?}"),
        }
    }
}

#[test]
fn paired_reference_limit_is_cumulative_and_charged_before_reads() {
    let resolver = LimitsResolver::standard(4);
    let (budget, _cancellation, execution) = make_execution("ods-formula-paired-cell-limit");
    let error = evaluate_source(
        &format!("=CORREL({X};{Y})"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("zero reference-cell budget must refuse before reads");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let resolver = LimitsResolver::standard(4);
    let (_budget, _cancellation, execution) = make_execution("ods-formula-paired-boundary");
    let result = evaluate_source(
        &format!("=COVAR({X};{Y})"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(8),
    )
    .expect("two four-cell arrays fit the exact cumulative boundary");
    assert_number(result, 2.5, "paired exact reference boundary");
    assert_eq!(resolver.reads(), 8);

    let resolver = LimitsResolver::standard(4);
    let (_budget, _cancellation, execution) = make_execution("ods-formula-paired-under-boundary");
    let error = evaluate_source(
        &format!("=COVAR({X};{Y})"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(7),
    )
    .expect_err("one cell below the pair total must refuse");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(
        resolver.reads(),
        7,
        "the cumulative cap is checked before the eighth read"
    );
}

#[test]
fn paired_list_and_three_dimensional_shape_refusals_are_read_free() {
    let resolver = LimitsResolver::standard(4);
    let (_budget, _cancellation, execution) = make_execution("ods-formula-paired-shapes");
    for source in [
        "=CORREL([.A1:.A2]~[.A3:.A4];[.B1:.B4])",
        "=COVAR([.A1:.A4];[.B1:.A2]~[.B3:.B4])",
        "=SLOPE([Main.B1:Archive.B2];[Main.A1:Archive.A2])",
        "=FORECAST(5;[.B1:.B4];[Main.A1:Archive.A2])",
    ] {
        let error = evaluate_source(source, &resolver, &execution, &Limits::default())
            .expect("shape refusal is a formula value");
        assert_eq!(error, LimitResult::Error(ScalarError::Value), "{source}");
    }
    assert_eq!(resolver.reads(), 0, "known rejected descriptors are unread");
}

#[test]
fn paired_shape_and_array_storage_limits_fail_atomically() {
    let resolver = LimitsResolver::standard(4);
    let (_budget, _cancellation, execution) = make_execution("ods-formula-paired-array-limit");
    let error = evaluate_source(
        "=CORREL({1;2;3};{1;2;3})",
        &resolver,
        &execution,
        &Limits::default().with_max_array_cells(2),
    )
    .expect_err("array-cell limit must reject before pair reduction");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);

    let error = evaluate_source(
        "=CORREL({1;2;3};{1;2;3})",
        &resolver,
        &execution,
        &Limits::default().with_max_storage_bytes(0),
    )
    .expect_err("zero storage must reject pair state");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn paired_work_text_and_provider_failures_are_typed_and_refunded() {
    let resolver = LimitsResolver::standard(4);
    let (budget, _cancellation, execution) = make_execution("ods-formula-paired-work-limit");
    let error = evaluate_source(
        &format!("=CORREL({X};{Y})"),
        &resolver,
        &execution,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("zero work must refuse before paired reads");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work
    ));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let error = evaluate_source(
        &format!("=IFERROR(CORREL({X};{Y});7)"),
        &resolver,
        &execution,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("IFERROR must not catch a typed paired resource refusal");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work
    ));
    assert_eq!(resolver.reads(), 0);

    let mut resolver = LimitsResolver::standard(4);
    resolver.set(1, 0, FixtureCell::Text("borrowed-text".to_owned()));
    let (budget, _cancellation, execution) = make_execution("ods-formula-paired-text-limit");
    let error = evaluate_source(
        &format!("=CORREL({X};{Y})"),
        &resolver,
        &execution,
        &Limits::default().with_max_text_bytes(1),
    )
    .expect_err("ignored paired text still consumes its text budget");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory
    ));
    assert_eq!(budget.used(Resource::Memory), 0);

    let resolver = LimitsResolver::standard(4).with_failure_at_read(8);
    let (budget, _cancellation, execution) = make_execution("ods-formula-paired-provider");
    let error = evaluate_source(
        &format!("=CORREL({X};{Y})"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(8),
    )
    .expect_err("provider failures must remain typed");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), 8);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn paired_formula_error_does_not_short_circuit_a_later_typed_read() {
    let mut resolver = LimitsResolver::standard(4);
    resolver.set(0, 0, FixtureCell::Error(ScalarError::NotAvailable));
    let resolver = resolver.with_failure_at_read(8);
    let (_budget, _cancellation, execution) = make_execution("ods-formula-paired-error-provider");
    let error = evaluate_source(
        &format!("=CORREL({X};{Y})"),
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("later typed provider failure supersedes retained formula error");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), 8);
}

#[test]
fn paired_cancellation_and_source_fences_surround_publication() {
    let (budget, cancellation, execution) = make_execution("ods-formula-paired-cancel");
    let resolver = LimitsResolver::standard(4).with_cancel_after_read(&cancellation);
    let error = evaluate_source(
        &format!("=SLOPE({Y};{X})"),
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("cancellation after the first paired read must fence publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);

    let mut resolver = LimitsResolver::standard(4);
    let expected = SourceVersion::new(0x5041_4952, 0);
    let observed = SourceVersion::new(0x5041_4952, 1);
    resolver.set_source_versions(expected, observed);
    let (budget, _cancellation, execution) = make_execution("ods-formula-paired-source");
    let error = evaluate_source(
        &format!("=COVAR({X};{Y})"),
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("source change must fence paired publication");
    assert!(matches!(
        error,
        EvaluationFailure::SourceChanged {
            expected: got_expected,
            observed: got_observed
        } if got_expected == expected && got_observed == observed
    ));
    assert!(resolver.reads() > 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn forecast_query_reference_lists_refuse_before_data_reads_and_query_errors_continue() {
    let resolver = LimitsResolver::standard(4);
    let (_budget, _cancellation, execution) = make_execution("ods-formula-paired-query-list");
    let error = evaluate_source(
        "=FORECAST([.G1:.G2]~[.G3:.G4];[.B1:.B4];[.A1:.A4])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("query list refusal is a formula value");
    assert_eq!(error, LimitResult::Error(ScalarError::Value));
    assert_eq!(resolver.reads(), 0, "query list gate precedes fit reads");

    let resolver = LimitsResolver::standard(4).with_failure_at_read(1);
    let (_budget, _cancellation, execution) = make_execution("ods-formula-paired-query-error");
    let result = evaluate_source(
        &format!("=FORECAST(\"bad\";{Y};{X})"),
        &LimitsResolver::standard(4),
        &execution,
        &Limits::default(),
    )
    .expect("malformed query is a formula value after the fit scan");
    assert_eq!(result, LimitResult::Error(ScalarError::Value));

    let error = evaluate_source(
        &format!("=FORECAST(\"bad\";{Y};{X})"),
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("typed fit failure supersedes malformed query conversion");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), 1);
}

#[test]
fn forecast_projected_queries_reuse_one_invariant_fit() {
    let resolver = LimitsResolver::standard(4);
    let (_budget, _cancellation, execution) = make_execution("ods-formula-paired-query-cache");
    let expression = parse("=IF(TRUE();FORECAST([.G1:.G2];[.B1:.B4];[.A1:.A4]);0)");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("projected FORECAST query");
    assert_array_shape_and_numbers(&result, 2, 1, &[11.0, 13.0], "projected FORECAST");
    assert_eq!(
        resolver.reads(),
        10,
        "fit data is read once for both queries"
    );
}

#[test]
fn paired_scalar_fit_streams_large_references_with_fixed_state() {
    let rows = 20_000;
    let resolver = LimitsResolver::standard(rows);
    let (_budget, _cancellation, execution) = make_execution("ods-formula-paired-large-stream");
    let limits = Limits::default()
        .with_max_reference_cells(rows * 2)
        .with_max_array_cells(1)
        .with_max_storage_bytes(64 * 1024)
        .with_max_steps(100_000_000);
    let result = evaluate_source(
        &format!("=CORREL([.A1:.A{rows}];[.B1:.B{rows}])"),
        &resolver,
        &execution,
        &limits,
    )
    .expect("fixed paired state should stream large references");
    assert_number(result, 1.0, "large paired correlation");
    assert_eq!(resolver.reads(), rows * 2);
}
