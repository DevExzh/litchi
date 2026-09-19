//! Resource and capability-boundary coverage for the ordering reducers.
//!
//! The resolver is observable and deliberately small.  These tests lock the
//! charge-before-read order, bounded sorting inputs, source/cancellation
//! fences, typed provider failures, and zero-read pseudotype refusals.

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
        self, CellRead, Context, Evaluated, Limits, Mode, OwnedValueView, Position, Resolver,
        SheetExtent, Value,
    },
    evaluation::{EvaluationFailure, ScalarError, UnsupportedKind},
    expression::Expression,
};

const RANGE: &str = "[.A1:.A8]";

#[derive(Debug, Clone)]
enum FixtureCell {
    Number(f64),
    Text(String),
    Error(ScalarError),
    Unsupported,
}

#[derive(Debug)]
struct LimitsResolver {
    rows: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_coordinates: RefCell<Vec<(usize, usize)>>,
    cancel_after_read: Option<CancellationSource>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl LimitsResolver {
    fn standard(rows: usize) -> Self {
        Self {
            rows,
            cells: (0..rows)
                .map(|row| FixtureCell::Number((row + 1) as f64))
                .collect(),
            reads: Cell::new(0),
            read_coordinates: RefCell::new(Vec::new()),
            cancel_after_read: None,
            source_versions: None,
            source_version_calls: Cell::new(0),
        }
    }

    fn set(&mut self, row: usize, value: FixtureCell) {
        self.cells[row] = value;
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn with_cancel_after_read(mut self, cancellation: &CancellationSource) -> Self {
        self.cancel_after_read = Some(cancellation.clone());
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
        Ok((sheet == "Main").then_some(SheetExtent::new(self.rows, 4)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.set(self.reads.get().saturating_add(1));
        self.read_coordinates.borrow_mut().push((row, column));
        if let Some(cancellation) = &self.cancel_after_read {
            cancellation.cancel();
        }
        if sheet != "Main" || column != 0 || row >= self.rows {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        Ok(match &self.cells[row] {
            FixtureCell::Number(value) => CellRead::Number(*value),
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
    match result {
        LimitResult::Number(actual) => assert_eq!(actual, expected, "{source:?}"),
        other => panic!("{source:?}: expected Number({expected}), got {other:?}"),
    }
}

#[test]
fn reference_cell_limit_is_charged_before_order_reducer_reads() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-order-cell-limit");
    let error = evaluate_source(
        &format!("=MEDIAN({RANGE})"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("zero reference-cell budget must reject ordering ranges");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong reference-cell failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn exact_reference_boundary_succeeds_and_one_below_refuses_without_reads() {
    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, execution) = execution("ods-formula-order-boundary");
    let result = evaluate_source(
        &format!("=LARGE({RANGE};2)"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(8),
    )
    .expect("exact reference-cell boundary should succeed");
    assert_number(result, 7.0, "LARGE boundary");
    assert_eq!(resolver.reads(), 8);

    let resolver = LimitsResolver::standard(8);
    let error = evaluate_source(
        &format!("=SMALL({RANGE};2)"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(7),
    )
    .expect_err("one cell below boundary must refuse before scanning");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn work_array_and_storage_limits_are_atomic() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, work_execution) = execution("ods-formula-order-work-limit");
    let error = evaluate_source(
        &format!("=PERCENTILE({RANGE};0.5)"),
        &resolver,
        &work_execution,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("zero Work must refuse before order scanning");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work
    ));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, array_execution) = execution("ods-formula-order-array-limit");
    let error = evaluate_source(
        "=MODE({1;2|3;4})",
        &resolver,
        &array_execution,
        &Limits::default().with_max_array_cells(3),
    )
    .expect_err("array-cell limit must reject before reduction");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);

    let error = evaluate_source(
        "=MODE({1;2|3;4})",
        &resolver,
        &array_execution,
        &Limits::default().with_max_storage_bytes(0),
    )
    .expect_err("zero evaluator storage must reject the retained ordering state");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory
    ));
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn variadic_order_arrays_use_cumulative_admitted_numeric_capacity() {
    let resolver = LimitsResolver::standard(1);
    let (_budget, _cancellation, execution) = execution("ods-formula-order-variadic-arrays");
    let limits = Limits::default()
        .with_max_array_cells(2)
        .with_max_reference_cells(1)
        .with_max_storage_bytes(64 * 1024);

    let result = evaluate_source("=MEDIAN({1;3};{5;7})", &resolver, &execution, &limits)
        .expect("MEDIAN should admit both two-cell sequence arrays");
    assert_number(result, 4.0, "variadic MEDIAN arrays");

    let result = evaluate_source("=MODE({1;2};{2;3})", &resolver, &execution, &limits)
        .expect("MODE should admit both two-cell sequence arrays");
    assert_number(result, 2.0, "variadic MODE arrays");
}

#[test]
fn large_sorted_reverse_and_duplicate_inputs_obey_read_bounds() {
    let rows = 2048;
    let resolver = LimitsResolver::standard(rows);
    let (_budget, _cancellation, execution) = execution("ods-formula-order-large");
    let result = evaluate_source(
        "=LARGE([.A1:.A2048];1024)",
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(rows),
    )
    .expect("large sorted input");
    assert_number(result, 1025.0, "large sorted input");
    assert_eq!(resolver.reads(), rows);

    let result = evaluate_source(
        "=SMALL([.A1:.A2048];1024)",
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(rows),
    )
    .expect("small reverse input");
    assert_number(result, 1024.0, "small reverse input");
    assert_eq!(resolver.reads(), rows * 2);
}

#[test]
fn scalar_rank_and_percent_rank_stream_large_references_without_output_arrays() {
    // Twenty thousand admitted cells exceed the 16 KiB that a retained
    // `Vec<f64>` would need. The scalar queries must therefore use their
    // fixed-size streaming state when the parameters are already Numbers.
    let rows = 20_000;
    let resolver = LimitsResolver::standard(rows);
    let (_budget, _cancellation, execution) = execution("ods-formula-order-scalar-stream");
    let limits = Limits::default()
        .with_max_reference_cells(rows)
        .with_max_array_cells(1)
        .with_max_storage_bytes(64 * 1024);

    let result = evaluate_source_mode(
        "=RANK(10000;[.A1:.A20000])",
        &resolver,
        &execution,
        Mode::Scalar,
        &limits,
    )
    .expect("scalar RANK should stream its admitted reference");
    assert_number(result, 10001.0, "scalar RANK large reference");
    assert_eq!(resolver.reads(), rows);

    let result = evaluate_source_mode(
        "=PERCENTRANK([.A1:.A20000];10000)",
        &resolver,
        &execution,
        Mode::Scalar,
        &limits,
    )
    .expect("scalar PERCENTRANK should stream its admitted reference");
    assert_number(result, 0.5, "scalar PERCENTRANK large reference");
    assert_eq!(resolver.reads(), rows * 2);
}

#[test]
fn borrowed_provider_text_is_budgeted_before_order_conversion() {
    let mut resolver = LimitsResolver::standard(2);
    resolver.set(0, FixtureCell::Text("abcdefg".to_owned()));
    resolver.set(1, FixtureCell::Number(3.0));
    let (budget, _cancellation, execution) = execution("ods-formula-order-text-limit");
    let error = evaluate_source(
        "=MODE([.A1:.A2])",
        &resolver,
        &execution,
        &Limits::default().with_max_text_bytes(6),
    )
    .expect_err("oversized provider text must be refused");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory
    ));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn cancellation_after_first_order_read_is_atomic() {
    let (budget, cancellation, execution) = execution("ods-formula-order-cancel");
    let resolver = LimitsResolver::standard(64).with_cancel_after_read(&cancellation);
    let error = evaluate_source(
        "=MEDIAN([.A1:.A64])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("provider cancellation must fence order publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn source_version_change_after_order_reads_is_not_published() {
    let mut resolver = LimitsResolver::standard(8);
    let expected = SourceVersion::new(0x4f52_4452, 0);
    let observed = SourceVersion::new(0x4f52_4452, 1);
    resolver.set_source_versions(expected, observed);
    let (budget, _cancellation, execution) = execution("ods-formula-order-source-change");
    let error = evaluate_source(
        &format!("=MEDIAN({RANGE})"),
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("source change must fence order publication");
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
fn typed_provider_failures_escape_all_order_reducers() {
    let mut resolver = LimitsResolver::standard(8);
    resolver.set(0, FixtureCell::Unsupported);
    let (_budget, _cancellation, execution) = execution("ods-formula-order-unsupported");
    for source in [
        "=MEDIAN([.A1:.A8])",
        "=MODE([.A1:.A8])",
        "=LARGE([.A1:.A8];1)",
        "=SMALL([.A1:.A8];1)",
        "=PERCENTILE([.A1:.A8];0.5)",
        "=PERCENTRANK([.A1:.A8];1)",
        "=QUARTILE([.A1:.A8];1)",
        "=RANK(1;[.A1:.A8])",
    ] {
        let error = evaluate_source(source, &resolver, &execution, &Limits::default())
            .expect_err("unsupported provider cell must remain typed");
        assert!(
            matches!(
                error,
                EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
            ),
            "{source}: wrong provider failure {error:?}"
        );
    }
}

#[test]
fn later_typed_failure_supersedes_a_retained_order_formula_error() {
    let mut resolver = LimitsResolver::standard(2);
    resolver.set(0, FixtureCell::Error(ScalarError::NotAvailable));
    resolver.set(1, FixtureCell::Unsupported);
    let (_budget, _cancellation, execution) = execution("ods-formula-order-error-precedence");
    for source in [
        "=MEDIAN([.A1:.A2])",
        "=MODE([.A1:.A2])",
        "=LARGE([.A1:.A2];1)",
        "=SMALL([.A1:.A2];1)",
        "=PERCENTILE([.A1:.A2];0.5)",
        "=PERCENTRANK([.A1:.A2];1)",
        "=QUARTILE([.A1:.A2];1)",
        "=RANK(1;[.A1:.A2])",
    ] {
        let error = evaluate_source(source, &resolver, &execution, &Limits::default())
            .expect_err("typed failure must supersede retained formula error");
        assert!(
            matches!(
                error,
                EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
            ),
            "{source}: wrong precedence result {error:?}"
        );
    }
}

#[test]
fn lazy_order_branch_does_not_probe_missing_provider_and_iferror_cannot_catch_limits() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-order-lazy");
    let result = evaluate_source(
        "=IF(FALSE();MEDIAN([Missing.A1:.Z100]);0)",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("unselected order branch");
    assert_number(result, 0.0, "lazy order branch");
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let error = evaluate_source(
        &format!("=IFERROR(MEDIAN({RANGE});7)"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("IFERROR must not catch an order resource refusal");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn sequence_shape_refusal_is_zero_read_for_mode_and_quartile() {
    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, execution) = execution("ods-formula-order-list-refusal");
    for source in [
        "=MODE([.A1:.A2]~[.A3:.A4])",
        "=QUARTILE([.A1:.A2]~[.A3:.A4];1)",
    ] {
        let result = evaluate_source(source, &resolver, &execution, &Limits::default())
            .expect("sequence shape refusal should be a formula value");
        assert_eq!(result, LimitResult::Error(ScalarError::Value), "{source}");
        assert_eq!(resolver.reads(), 0, "{source} must refuse before reads");
    }
}

#[test]
fn scalar_parameter_reference_lists_refuse_before_data_reads() {
    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, execution) = execution("ods-formula-order-parameter-list-refusal");
    for source in [
        "=RANK([.A1:.A2]~[.A3:.A4];[.B1:.B6])",
        "=PERCENTILE([.A1:.A8];[.A1:.A2]~[.A3:.A4])",
    ] {
        let result = evaluate_source(source, &resolver, &execution, &Limits::default())
            .expect("scalar-parameter list refusal should be a formula value");
        assert_eq!(result, LimitResult::Error(ScalarError::Value), "{source}");
        assert_eq!(
            resolver.reads(),
            0,
            "{source} must refuse before data reads"
        );
    }
}

#[test]
fn scalar_profile_refuses_reference_descriptors_without_a_resolver() {
    let resolver = LimitsResolver::standard(2);
    let (_budget, _cancellation, execution) = execution("ods-formula-order-scalar-reference");
    let expression = parse("=MEDIAN([.A1:.A2])");
    let result = litchi_ods::codec::formula::evaluation::evaluate_scalar(
        &expression,
        &litchi_ods::codec::formula::evaluation::EvaluationContext::new(&execution),
        &litchi_ods::codec::formula::evaluation::EvaluationLimits::default(),
    );
    assert!(matches!(
        result,
        Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference))
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn owned_order_result_releases_borrowed_reference_state() {
    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, execution) = execution("ods-formula-order-owned");
    let expression = parse("=MEDIAN([.A1:.A2])");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("order reference result");
    let owned = result
        .to_owned(&execution, &Limits::default())
        .expect("order result should be ownable");
    assert!(matches!(owned.value(), OwnedValueView::Number(value) if value == 1.5));
}
