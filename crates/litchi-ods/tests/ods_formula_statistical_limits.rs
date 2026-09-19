//! Resource, cancellation, source-freshness, lazy-branch, and result
//! ownership coverage for the OpenFormula statistical reducer family.
//!
//! Every refusal is checked at the public value-evaluator boundary. The
//! resolver records reads so a geometry or array limit must be charged before
//! provider access, and every failed path must release evaluator-owned memory.

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
    Empty,
    Number(f64),
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
            FixtureCell::Empty => CellRead::Empty,
            FixtureCell::Number(value) => CellRead::Number(*value),
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
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
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
    position: Position<'a>,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, position).with_mode(mode);
    value::evaluate(expression, resolver, &context, limits)
}

fn assert_number(result: LimitResult, expected: f64, source: &str) {
    match result {
        LimitResult::Number(actual) => assert_eq!(actual, expected, "{source:?}"),
        other => panic!("{source:?}: expected Number({expected}), got {other:?}"),
    }
}

#[test]
fn reference_cell_limit_is_charged_before_provider_reads() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-statistical-cell-limit");
    let error = evaluate_source(
        &format!("=COUNT({RANGE})"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("zero reference-cell budget must reject statistical ranges");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong reference-cell failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn reference_area_limit_is_atomic_and_read_free() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-statistical-area-limit");
    let error = evaluate_source(
        "=COUNT([.A1]~[.A2])",
        &resolver,
        &execution,
        &Limits::default().with_max_reference_areas(0),
    )
    .expect_err("zero reference-area budget must reject a list");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong reference-area failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn exact_reference_cell_boundary_succeeds_and_under_boundary_refuses() {
    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-cell-boundary");
    let result = evaluate_source(
        &format!("=COUNT({RANGE})"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(8),
    )
    .expect("the exact reference-cell boundary should succeed");
    assert_number(result, 8.0, "COUNT boundary");
    assert_eq!(resolver.reads(), 8);

    let resolver = LimitsResolver::standard(8);
    let error = evaluate_source(
        &format!("=COUNT({RANGE})"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(7),
    )
    .expect_err("one cell below the boundary must refuse before scanning");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn statistical_work_and_array_limits_are_typed_and_release_memory() {
    let mut resolver = LimitsResolver::standard(8);
    resolver.set(7, FixtureCell::Empty);
    let (budget, _cancellation, work_execution) = execution("ods-formula-statistical-work-limit");
    let error = evaluate_source(
        &format!("=AVERAGE({RANGE})"),
        &resolver,
        &work_execution,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("zero Work must refuse before statistical scanning");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong Work failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, array_execution) = execution("ods-formula-statistical-array-limit");
    let error = evaluate_source(
        "=AVERAGE({1;2|3;4})",
        &resolver,
        &array_execution,
        &Limits::default().with_max_array_cells(3),
    )
    .expect_err("array-cell limit must reject before reduction");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong array-cell failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn statistical_storage_limit_is_atomic() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-statistical-storage-limit");
    let error = evaluate_source(
        "=AVERAGEA({1;2|3;4})",
        &resolver,
        &execution,
        &Limits::default().with_max_storage_bytes(0),
    )
    .expect_err("zero evaluator storage must refuse array admission");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong storage failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn cancellation_after_the_first_statistical_read_is_atomic() {
    let (budget, cancellation, execution) = execution("ods-formula-statistical-cancel");
    let resolver = LimitsResolver::standard(64).with_cancel_after_read(&cancellation);
    let error = evaluate_source(
        "=MAX([.A1:.A64])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("provider cancellation must fence statistical publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn source_version_change_after_statistical_reads_is_not_published() {
    let mut resolver = LimitsResolver::standard(8);
    let expected = SourceVersion::new(0x5354_4154, 0);
    let observed = SourceVersion::new(0x5354_4154, 1);
    resolver.set_source_versions(expected, observed);
    let (budget, _cancellation, execution) = execution("ods-formula-statistical-source-change");
    let error = evaluate_source(
        &format!("=AVERAGE({RANGE})"),
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("source change must fence statistical publication");
    assert!(
        matches!(
            error,
            EvaluationFailure::SourceChanged {
                expected: got_expected,
                observed: got_observed
            } if got_expected == expected && got_observed == observed
        ),
        "wrong source-change failure: {error:?}"
    );
    assert!(resolver.reads() > 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn unsupported_provider_cells_remain_typed_failures() {
    let mut resolver = LimitsResolver::standard(8);
    resolver.set(0, FixtureCell::Unsupported);
    let (budget, _cancellation, execution) = execution("ods-formula-statistical-unsupported");
    let error = evaluate_source(
        &format!("=AVERAGE({RANGE})"),
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("unsupported provider cells must not become formula errors");
    assert!(
        matches!(
            error,
            EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
        ),
        "wrong provider failure: {error:?}"
    );
    assert!(resolver.reads() > 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn lazy_statistical_branches_do_not_probe_missing_sheets() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-statistical-lazy");
    let result = evaluate_source(
        "=IF(FALSE();AVERAGE([Missing.A1:.Z100]);0)",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("unselected statistical branch");
    assert_number(result, 0.0, "lazy statistical branch");
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn iferror_cannot_catch_a_statistical_resource_refusal() {
    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-iferror-limit");
    let error = evaluate_source(
        &format!("=IFERROR(AVERAGE({RANGE});7)"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("IFERROR must not catch a statistical resource refusal");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn statistical_result_ownership_releases_borrowed_inputs() {
    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-owned");
    let expression = parse("=AVERAGE([.A1:.A2])");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("statistical reference result");
    let owned = result
        .to_owned(&execution, &Limits::default())
        .expect("statistical result should be ownable");
    assert!(matches!(owned.value(), OwnedValueView::Number(value) if value == 1.5));
}

#[test]
fn deeply_nested_statistical_projection_is_bounded_by_work() {
    let mut nested = "1".to_owned();
    for _ in 0..192 {
        nested = format!("AVERAGE({nested})");
    }
    let source = format!("=IF({{TRUE();TRUE()}};{nested};0)");
    let resolver = LimitsResolver::standard(2);
    let (budget, _cancellation, execution) = execution("ods-formula-statistical-deep");
    let expression = parse(&source);
    let error = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default().with_max_steps(256),
    )
    .expect_err("deep statistical projection must be bounded");
    assert!(
        matches!(
            &error,
            EvaluationFailure::ResourceLimit(limit)
                if matches!(limit.resource, Resource::Work | Resource::Depth)
        ),
        "wrong deep-projection failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}
