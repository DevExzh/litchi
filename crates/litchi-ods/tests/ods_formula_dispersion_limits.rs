//! Resource and capability-boundary coverage for the dispersion reducers.
//!
//! The resolver is deliberately small and observable.  Each test checks that
//! geometry, array, text, work, cancellation, and source-freshness decisions
//! happen at the value-evaluator boundary and that typed provider failures are
//! not converted into formula errors by a dispersion reducer.

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

fn assert_close(result: LimitResult, expected: f64, source: &str) {
    let LimitResult::Number(actual) = result else {
        panic!("{source:?}: expected Number({expected}), got {result:?}");
    };
    let tolerance = expected.abs().max(actual.abs()).max(1.0) * 64.0 * f64::EPSILON;
    assert!(
        (actual - expected).abs() <= tolerance,
        "{source:?}: {actual} != {expected} (tolerance {tolerance})"
    );
}

#[test]
fn reference_limit_is_charged_before_dispersion_reads() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-dispersion-cell-limit");
    let error = evaluate_source(
        &format!("=VARP({RANGE})"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("zero reference-cell budget must reject dispersion ranges");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong reference-cell failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn array_and_storage_limits_refuse_before_dispersion_reduction() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-dispersion-array-limit");
    let error = evaluate_source(
        "=VAR({1;2|3;4})",
        &resolver,
        &execution,
        &Limits::default().with_max_array_cells(3),
    )
    .expect_err("array-cell limit must reject before reduction");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong array-cell failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 0);

    let error = evaluate_source(
        "=VARA({1;2|3;4})",
        &resolver,
        &execution,
        &Limits::default().with_max_storage_bytes(0),
    )
    .expect_err("storage limit must reject array admission");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong storage failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn provider_text_is_borrowed_and_charged_before_a_variant_coercion() {
    let mut resolver = LimitsResolver::standard(2);
    resolver.set(0, FixtureCell::Text("abcdefg".to_owned()));
    resolver.set(1, FixtureCell::Number(3.0));
    let (budget, _cancellation, execution) = execution("ods-formula-dispersion-text-limit");
    let error = evaluate_source(
        "=VARA([.A1:.A2])",
        &resolver,
        &execution,
        &Limits::default().with_max_text_bytes(6),
    )
    .expect_err("oversized provider text must be refused");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong text-limit failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);

    let expression = parse("=VARPA([.A1:.A2])");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("provider text should be borrowed for a successful A reducer");
    let owned = result
        .to_owned(&execution, &Limits::default())
        .expect("dispersion result should be ownable");
    assert!(matches!(owned.value(), OwnedValueView::Number(_)));
}

#[test]
fn cancellation_after_first_dispersion_read_is_atomic() {
    let (budget, cancellation, execution) = execution("ods-formula-dispersion-cancel");
    let resolver = LimitsResolver::standard(64).with_cancel_after_read(&cancellation);
    let error = evaluate_source(
        "=STDEVP([.A1:.A64])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("provider cancellation must fence dispersion publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn source_version_change_after_dispersion_reads_is_not_published() {
    let mut resolver = LimitsResolver::standard(8);
    let expected = SourceVersion::new(0x4453_5052, 0);
    let observed = SourceVersion::new(0x4453_5052, 1);
    resolver.set_source_versions(expected, observed);
    let (budget, _cancellation, execution) = execution("ods-formula-dispersion-source-change");
    let error = evaluate_source(
        &format!("=VAR({RANGE})"),
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("source change must fence dispersion publication");
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
fn typed_provider_failures_escape_all_dispersion_variants() {
    let mut resolver = LimitsResolver::standard(8);
    resolver.set(0, FixtureCell::Unsupported);
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-unsupported");
    for function in [
        "VAR", "VARA", "VARP", "VARPA", "STDEV", "STDEVA", "STDEVP", "STDEVPA",
    ] {
        let source = format!("={function}({RANGE})");
        let error = evaluate_source(&source, &resolver, &execution, &Limits::default())
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
fn later_typed_failure_supersedes_a_retained_formula_error() {
    let mut resolver = LimitsResolver::standard(2);
    resolver.set(0, FixtureCell::Error(ScalarError::NotAvailable));
    resolver.set(1, FixtureCell::Unsupported);
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-error-precedence");
    for function in [
        "VAR", "VARA", "VARP", "VARPA", "STDEV", "STDEVA", "STDEVP", "STDEVPA",
    ] {
        let source = format!("={function}([.A1:.A2])");
        let error = evaluate_source(&source, &resolver, &execution, &Limits::default())
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
fn unselected_dispersion_branch_does_not_probe_provider() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-dispersion-lazy");
    let result = evaluate_source(
        "=IF(FALSE();VARP([Missing.A1:.Z100]);0)",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("unselected dispersion branch");
    assert_number(result, 0.0, "lazy dispersion branch");
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn iferror_cannot_catch_a_dispersion_resource_refusal() {
    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-iferror-limit");
    let error = evaluate_source(
        &format!("=IFERROR(STDEVP({RANGE});7)"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("IFERROR must not catch a dispersion resource refusal");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn direct_dispersion_shape_refusal_has_no_resolver_reads() {
    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-direct");
    let result = evaluate_source("=VARP(1;2;3)", &resolver, &execution, &Limits::default())
        .expect("direct dispersion arguments must not consult a resolver");
    assert_close(result, 2.0 / 3.0, "=VARP(1;2;3)");
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn number_sequence_reference_list_refusal_has_zero_reads() {
    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-list-refusal");
    for function in ["VAR", "VARP", "STDEVP"] {
        let source = format!("={function}([.A1:.A2]~[.A3])");
        let result = evaluate_source(&source, &resolver, &execution, &Limits::default())
            .expect("reference-list shape refusal should be a formula value");
        assert_eq!(result, LimitResult::Error(ScalarError::Value), "{source}");
        assert_eq!(
            resolver.reads(),
            0,
            "{source} must refuse before resolver reads"
        );
    }
}
