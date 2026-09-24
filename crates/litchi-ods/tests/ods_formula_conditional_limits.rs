//! Resource, cancellation, source-freshness, and ownership coverage for the
//! OpenFormula conditional aggregate family.
//!
//! Conditional aggregates stream reference cells through the caller's value
//! evaluator.  The resolver below records every provider read and can fence a
//! read, source-version sample, or selected branch so a failed evaluation must
//! leave no partially published result or retained evaluator memory.

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
    evaluation::{EvaluationFailure, ScalarError},
    expression::Expression,
};

const CRITERIA: &str = "[.A1:.A8]";
const VALUES: &str = "[.C1:.C8]";

#[derive(Debug, Clone)]
enum FixtureCell {
    Empty,
    Number(f64),
    Unsupported,
}

#[derive(Debug)]
struct LimitsResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_coordinates: RefCell<Vec<(usize, usize)>>,
    cancel_after_read: Option<CancellationSource>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl LimitsResolver {
    fn standard(rows: usize) -> Self {
        let mut resolver = Self {
            rows,
            columns: 6,
            cells: (0..rows * 6).map(|_| FixtureCell::Empty).collect(),
            reads: Cell::new(0),
            read_coordinates: RefCell::new(Vec::new()),
            cancel_after_read: None,
            source_versions: None,
            source_version_calls: Cell::new(0),
        };
        for row in 0..rows {
            resolver.set(
                row,
                0,
                FixtureCell::Number(if row % 2 == 0 { 1.0 } else { 2.0 }),
            );
            resolver.set(row, 2, FixtureCell::Number((row + 1) as f64));
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

    fn with_cancel_after_read(mut self, cancellation: &CancellationSource) -> Self {
        self.cancel_after_read = Some(cancellation.clone());
        self
    }

    fn set_source_versions(&mut self, expected: SourceVersion, observed: SourceVersion) {
        self.source_versions = Some((expected, observed));
        self.source_version_calls.set(0);
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn read_coordinates(&self) -> Vec<(usize, usize)> {
        self.read_coordinates.borrow().clone()
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
        self.read_coordinates.borrow_mut().push((row, column));
        if let Some(cancellation) = &self.cancel_after_read {
            cancellation.cancel();
        }
        if sheet != "Main" {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        let Some(index) = row
            .checked_mul(self.columns)
            .and_then(|index| index.checked_add(column))
        else {
            return Ok(CellRead::Empty);
        };
        Ok(match self.cells.get(index) {
            None | Some(FixtureCell::Empty) => CellRead::Empty,
            Some(FixtureCell::Number(value)) => CellRead::Number(*value),
            Some(FixtureCell::Unsupported) => CellRead::Unsupported,
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
        NonZeroU64::new(1_024).expect("one in-flight KiB"),
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

// This helper returns a lifetime-free observation.  The parsed expression is
// local, so returning `Evaluated` here would let its expression metadata (and
// any borrowed resolver text) outlive the inputs.  The ownership boundary is
// tested separately through `evaluate_expression` and `Evaluated::to_owned`.
fn evaluate(
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

fn evaluate_expression<'a>(
    expression: &'a Expression,
    resolver: &'a LimitsResolver,
    execution: &ExecutionContext,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    value::evaluate(expression, resolver, &context, limits)
}

fn assert_number(result: &LimitResult, expected: f64, source: &str) {
    match result {
        LimitResult::Number(actual) => assert_eq!(*actual, expected, "{source:?}"),
        other => panic!("{source:?}: expected Number({expected}), got {other:?}"),
    }
}

fn assert_evaluated_number(result: &Evaluated<'_>, expected: f64, source: &str) {
    match result.value() {
        Value::Number(actual) => assert_eq!(actual, expected, "{source:?}"),
        other => panic!("{source:?}: expected Number({expected}), got {other:?}"),
    }
}

#[test]
fn conditional_reference_cell_limit_refuses_before_provider_reads() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-conditional-cell-limit");
    let source = format!("=SUMIF({CRITERIA};1;{VALUES})");
    let error = evaluate(
        &source,
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("zero reference-cell budget must reject conditional ranges");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong reference-cell failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn conditional_work_limit_is_typed_and_does_not_publish_memory() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-conditional-work-limit");
    let source = format!("=COUNTIFS({CRITERIA};1)");
    let error = evaluate(
        &source,
        &resolver,
        &execution,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("zero Work must reject conditional evaluation");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong Work failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn conditional_matcher_storage_limit_fails_before_range_reads() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-conditional-storage-limit");
    let source = format!("=COUNTIFS({CRITERIA};1;{CRITERIA};2)");
    let error = evaluate(
        &source,
        &resolver,
        &execution,
        &Limits::default().with_max_storage_bytes(0),
    )
    .expect_err("zero evaluator storage must reject conditional matcher planning");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong conditional storage failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn provider_cancellation_after_the_first_conditional_read_is_atomic() {
    let (budget, cancellation, execution) = execution("ods-formula-conditional-cancel");
    let resolver = LimitsResolver::standard(8).with_cancel_after_read(&cancellation);
    let source = format!("=SUMIFS({VALUES};{CRITERIA};1)");
    let error = evaluate(&source, &resolver, &execution, &Limits::default())
        .expect_err("provider cancellation must fence conditional publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn unsupported_selected_provider_cells_remain_typed_failures() {
    let mut resolver = LimitsResolver::standard(8);
    resolver.set(0, 2, FixtureCell::Unsupported);
    let (budget, _cancellation, execution) = execution("ods-formula-conditional-unsupported-cell");
    let source = format!("=SUMIF({CRITERIA};1;{VALUES})");
    let error = evaluate(&source, &resolver, &execution, &Limits::default())
        .expect_err("unsupported selected destination must stay a typed failure");
    assert!(
        matches!(
            error,
            EvaluationFailure::Unsupported(
                litchi_ods::codec::formula::evaluation::UnsupportedKind::CellValue
            )
        ),
        "wrong unsupported-cell failure: {error:?}"
    );
    assert!(resolver.reads() > 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn source_version_change_after_conditional_reads_is_not_published() {
    let mut resolver = LimitsResolver::standard(8);
    let expected = SourceVersion::new(0x434f4e44, 0);
    let observed = SourceVersion::new(0x434f4e44, 1);
    resolver.set_source_versions(expected, observed);
    let (budget, _cancellation, execution) = execution("ods-formula-conditional-source-change");
    let source = format!("=AVERAGEIF({VALUES};\">0\")");
    let error = evaluate(&source, &resolver, &execution, &Limits::default())
        .expect_err("source change must fence conditional publication");
    assert!(
        matches!(error, EvaluationFailure::SourceChanged { expected: got_expected, observed: got_observed } if got_expected == expected && got_observed == observed),
        "wrong source-change failure: {error:?}"
    );
    assert!(resolver.reads() > 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn unselected_conditional_branch_does_not_probe_the_provider() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-conditional-lazy-branch");
    let source = format!("=IF(FALSE();SUMIF({CRITERIA};1;{VALUES});0)");
    let result = evaluate(&source, &resolver, &execution, &Limits::default())
        .expect("unselected conditional branch");
    assert_number(&result, 0.0, &source);
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn reference_list_area_limit_is_checked_before_conditional_reads() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-conditional-area-limit");
    let source = "=SUMIF(([.A1:.A4]~[.A5:.A8]);1)";
    let error = evaluate(
        source,
        &resolver,
        &execution,
        &Limits::default().with_max_reference_areas(1),
    )
    .expect_err("reference-list area limit must refuse before reads");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong area-limit failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn conditional_provider_reads_scale_with_two_ranges_and_stop_at_the_cap() {
    let source = format!("=SUMIF({CRITERIA};1;{VALUES})+COUNTIF({CRITERIA};1)");
    let baseline_resolver = LimitsResolver::standard(32);
    let (baseline_budget, _baseline_cancellation, baseline_execution) =
        execution("ods-formula-conditional-read-baseline");
    let baseline = evaluate(
        &source,
        &baseline_resolver,
        &baseline_execution,
        &Limits::default(),
    )
    .expect("conditional read baseline");
    assert!(matches!(baseline, LimitResult::Number(_)));
    let successful_reads = baseline_resolver.reads();
    assert!(successful_reads > 0);
    assert_eq!(baseline_budget.used(Resource::Memory), 0);

    let cap = successful_reads.saturating_sub(1);
    let resolver = LimitsResolver::standard(32);
    let (budget, _cancellation, execution) = execution("ods-formula-conditional-read-cap");
    let error = evaluate(
        &source,
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(cap),
    )
    .expect_err("cumulative conditional reads must stop at the cap");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong cumulative reference failure: {error:?}"
    );
    assert_eq!(resolver.reads(), cap);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn owned_conditional_number_survives_expression_resolver_and_execution_drop() {
    let (budget, _cancellation, execution) = execution("ods-formula-conditional-owned-number");
    let owned = {
        let source = format!("=SUMIF({CRITERIA};1;{VALUES})");
        let expression = parse(&source);
        let resolver = LimitsResolver::standard(8);
        let borrowed = evaluate_expression(&expression, &resolver, &execution, &Limits::default())
            .expect("conditional number");
        assert_evaluated_number(&borrowed, 16.0, &source);
        borrowed
            .to_owned(&execution, &Limits::default())
            .expect("conditional number ownership")
    };
    drop(execution);
    assert!(matches!(owned.value(), OwnedValueView::Number(16.0)));
    assert_eq!(owned.reserved_storage_bytes(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
    drop(owned);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn provider_read_coordinates_remain_in_reference_order() {
    let resolver = LimitsResolver::standard(4);
    let (_budget, _cancellation, execution) = execution("ods-formula-conditional-read-order");
    let source = "=SUMIFS([.C1:.C4];[.A1:.A4];1)";
    let result = evaluate(source, &resolver, &execution, &Limits::default())
        .expect("conditional read order");
    assert_number(&result, 4.0, source);
    assert!(!resolver.read_coordinates().is_empty());
}
