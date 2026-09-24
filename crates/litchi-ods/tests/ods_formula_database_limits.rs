//! Bounded limits, cancellation, and ownership coverage for ODF database
//! functions.
//!
//! The ordinary database semantics live in `ods_formula_database_evaluation`.
//! This file keeps its provider deliberately observable so resource admission,
//! lazy branches, projected scalar caching, and the DGET ownership boundary can
//! be checked independently.

use std::cell::{Cell, RefCell};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Evaluated, Limits as ValueLimits, Mode, OwnedValueView, Position,
        Resolver, SheetExtent, Value,
    },
    evaluation::{EvaluationFailure, ScalarError},
    expression::Expression,
};

const DATABASE: &str = "[.A1:.D4]";
const CRITERIA: &str = "[.F1:.F2]";
const ROWS: usize = 4;
const COLUMNS: usize = 7;
const DATABASE_GEOMETRY: usize = 16;

#[derive(Debug)]
enum FixtureCell {
    Empty,
    Number(f64),
    Text(String),
    Unsupported,
}

#[derive(Debug)]
struct DatabaseLimitsResolver {
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_coordinates: RefCell<Vec<(usize, usize)>>,
    cancel_after_read: Option<CancellationSource>,
}

impl DatabaseLimitsResolver {
    fn standard(criteria: &str) -> Self {
        let mut resolver = Self {
            cells: (0..ROWS * COLUMNS).map(|_| FixtureCell::Empty).collect(),
            reads: Cell::new(0),
            read_coordinates: RefCell::new(Vec::new()),
            cancel_after_read: None,
        };
        for (column, header) in ["Name", "Region", "Amount", "Unused"]
            .into_iter()
            .enumerate()
        {
            resolver.set(0, column, FixtureCell::Text(header.to_owned()));
        }
        resolver.record(1, "Alice", "East", 10.0, "private");
        resolver.record(2, "Bob", "West", 20.0, "private");
        resolver.record(3, "Carol", "East", 30.0, "private");
        resolver.set(0, 5, FixtureCell::Text("Region".to_owned()));
        resolver.set(1, 5, FixtureCell::Text(criteria.to_owned()));
        resolver
    }

    fn set(&mut self, row: usize, column: usize, value: FixtureCell) {
        let index = row
            .checked_mul(COLUMNS)
            .and_then(|index| index.checked_add(column))
            .expect("fixture coordinate fits");
        self.cells[index] = value;
    }

    fn record(&mut self, row: usize, name: &str, region: &str, amount: f64, unused: &str) {
        self.set(row, 0, FixtureCell::Text(name.to_owned()));
        self.set(row, 1, FixtureCell::Text(region.to_owned()));
        self.set(row, 2, FixtureCell::Number(amount));
        self.set(row, 3, FixtureCell::Text(unused.to_owned()));
    }

    fn with_cancel_after_read(mut self, cancellation: &CancellationSource) -> Self {
        self.cancel_after_read = Some(cancellation.clone());
        self
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn read_coordinates(&self) -> Vec<(usize, usize)> {
        self.read_coordinates.borrow().clone()
    }
}

impl Resolver for DatabaseLimitsResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(SheetExtent::new(ROWS, COLUMNS)))
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
            .checked_mul(COLUMNS)
            .and_then(|index| index.checked_add(column))
        else {
            return Ok(CellRead::Empty);
        };
        Ok(match self.cells.get(index) {
            None | Some(FixtureCell::Empty) => CellRead::Empty,
            Some(FixtureCell::Number(value)) => CellRead::Number(*value),
            Some(FixtureCell::Text(value)) => CellRead::Text(value.as_str()),
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
}

fn make_execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroU64::new(1 << 20).expect("nonzero in-flight byte limit"),
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

fn evaluate<'a>(
    expression: &'a Expression,
    resolver: &'a DatabaseLimitsResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &ValueLimits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    value::evaluate(expression, resolver, &context, limits)
}

fn assert_number(result: &Evaluated<'_>, expected: f64, source: &str) {
    match result.value() {
        Value::Number(actual) => assert_eq!(actual, expected, "{source:?}"),
        other => panic!("{source:?}: expected Number({expected}), got {other:?}"),
    }
}

#[test]
fn database_zero_work_and_storage_limits_are_typed_and_atomic() {
    let expression = parse(&format!("=DSUM({DATABASE};\"Amount\";{CRITERIA})"));
    let resolver = DatabaseLimitsResolver::standard("East");

    let (budget, _cancellation, execution) = make_execution("ods-formula-database-work-limit");
    let error = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &ValueLimits::default().with_max_steps(0),
    )
    .expect_err("zero Work must refuse database evaluation");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong database Work failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, execution) = make_execution("ods-formula-database-memory-limit");
    let error = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &ValueLimits::default().with_max_storage_bytes(0),
    )
    .expect_err("zero storage must refuse database evaluation");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong database storage failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn database_reference_read_admission_is_cumulative_and_atomic() {
    let source =
        format!("=DSUM({DATABASE};\"Amount\";{CRITERIA})+DCOUNT({DATABASE};\"Name\";{CRITERIA})");
    let expression = parse(&source);
    let successful_reads = {
        let resolver = DatabaseLimitsResolver::standard("East");
        let (budget, _cancellation, execution) =
            make_execution("ods-formula-database-cumulative-reference-baseline");
        let result = evaluate(
            &expression,
            &resolver,
            &execution,
            Mode::Scalar,
            &ValueLimits::default(),
        )
        .expect("two independent database calls should fit the default reference budget");
        assert_number(&result, 40.0, &source);
        let reads = resolver.reads();
        assert!(
            reads >= DATABASE_GEOMETRY,
            "successful database evaluation must admit its physical geometry: {reads}"
        );
        drop(result);
        assert_eq!(budget.used(Resource::Memory), 0);
        reads
    };
    let max_reference_cells = successful_reads
        .checked_sub(1)
        .expect("successful evaluation must perform at least one provider read");
    assert!(
        max_reference_cells >= DATABASE_GEOMETRY,
        "reference cap must leave room for one complete database geometry: {max_reference_cells}"
    );

    let resolver = DatabaseLimitsResolver::standard("East");
    let (budget, _cancellation, execution) =
        make_execution("ods-formula-database-cumulative-reference-limit");
    let error = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &ValueLimits::default().with_max_reference_cells(max_reference_cells),
    )
    .expect_err("cumulative database reads should exceed the reference cap");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong cumulative reference failure: {error:?}"
    );
    assert_eq!(
        resolver.reads(),
        max_reference_cells,
        "provider reads must stop exactly at the cumulative reference cap"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn false_criteria_clause_short_circuits_later_unsupported_field_reads() {
    let source = format!("=DSUM({DATABASE};\"Amount\";[.F1:.G2])");
    let expression = parse(&source);
    let mut resolver = DatabaseLimitsResolver::standard("Never");
    resolver.set(0, 6, FixtureCell::Text("Name".to_owned()));
    resolver.set(1, 6, FixtureCell::Text("Alice".to_owned()));
    resolver.set(1, 0, FixtureCell::Unsupported);
    let (budget, _cancellation, execution) =
        make_execution("ods-formula-database-criteria-short-circuit");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &ValueLimits::default(),
    )
    .expect("a false first criterion clause should avoid the unsupported field");
    assert_number(&result, 0.0, &source);
    assert!(
        !resolver
            .read_coordinates()
            .iter()
            .any(|&(row, column)| row > 0 && column == 0),
        "the later Name criterion must not read the unsupported record field"
    );
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn duplicate_criteria_fields_share_one_provider_read_per_database_row() {
    let source = format!("=DCOUNT({DATABASE};\"Amount\";[.F1:.G2])");
    let expression = parse(&source);
    let mut resolver = DatabaseLimitsResolver::standard("East");
    resolver.set(0, 6, FixtureCell::Text("Region".to_owned()));
    resolver.set(1, 6, FixtureCell::Text("East".to_owned()));
    let (budget, _cancellation, execution) =
        make_execution("ods-formula-database-criteria-duplicate-field");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &ValueLimits::default(),
    )
    .expect("duplicate criteria fields should evaluate");
    assert_number(&result, 2.0, &source);
    let coordinates = resolver.read_coordinates();
    for row in 1..ROWS {
        assert_eq!(
            coordinates
                .iter()
                .filter(|&&(read_row, column)| read_row == row && column == 1)
                .count(),
            1,
            "duplicate Region criteria must read one Region cell for row {row}: {coordinates:?}"
        );
    }
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn oversized_provider_text_is_refused_before_database_comparison() {
    let source = format!("=DSUM({DATABASE};\"Amount\";{CRITERIA})");
    for (label, row, column) in [("header", 0, 0), ("criterion", 1, 5)] {
        let expression = parse(&source);
        let mut resolver = DatabaseLimitsResolver::standard("East");
        resolver.set(row, column, FixtureCell::Text("x".repeat(7)));
        let (budget, _cancellation, execution) =
            make_execution("ods-formula-database-provider-text-limit");
        let error = evaluate(
            &expression,
            &resolver,
            &execution,
            Mode::Scalar,
            &ValueLimits::default().with_max_text_bytes(6),
        )
        .expect_err("oversized provider text must be refused");
        assert!(
            matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
            "wrong {label} text-limit failure: {error:?}"
        );
        assert!(
            !resolver
                .read_coordinates()
                .iter()
                .any(|&(read_row, read_column)| read_row > 0 && read_column < 4),
            "{label} text limit must refuse before reading database records"
        );
        assert_eq!(budget.used(Resource::Memory), 0);
    }
}

#[test]
fn database_provider_cancellation_stops_mid_scan_and_releases_memory() {
    let source = format!("=DSUM({DATABASE};\"Amount\";{CRITERIA})");
    let expression = parse(&source);
    let (budget, cancellation, execution) = make_execution("ods-formula-database-mid-read-cancel");
    let resolver = DatabaseLimitsResolver::standard("East").with_cancel_after_read(&cancellation);
    let error = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &ValueLimits::default(),
    )
    .expect_err("provider cancellation must fence database publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert!(budget.used(Resource::Work) > 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn unselected_database_branch_is_lazy_and_does_not_read_the_provider() {
    let source = format!("=IF(FALSE();DSUM({DATABASE};\"Amount\";{CRITERIA});0)");
    let expression = parse(&source);
    let resolver = DatabaseLimitsResolver::standard("East");
    let (budget, _cancellation, execution) = make_execution("ods-formula-database-lazy");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &ValueLimits::default(),
    )
    .expect("unselected database branch must not be resolved");
    assert_number(&result, 0.0, &source);
    assert_eq!(resolver.reads(), 0);
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn projected_scalar_database_result_is_cached_and_broadcast_with_shape() {
    let source = format!("=IF({{TRUE();FALSE();TRUE()}};DSUM({DATABASE};\"Amount\";{CRITERIA});0)");
    let expression = parse(&source);

    let baseline_source = format!("=DSUM({DATABASE};\"Amount\";{CRITERIA})");
    let baseline_expression = parse(&baseline_source);
    let baseline_reads = {
        let baseline_resolver = DatabaseLimitsResolver::standard("East");
        let (baseline_budget, _baseline_cancellation, baseline_execution) =
            make_execution("ods-formula-database-projected-cache-baseline");
        let baseline = evaluate(
            &baseline_expression,
            &baseline_resolver,
            &baseline_execution,
            Mode::Scalar,
            &ValueLimits::default(),
        )
        .expect("scalar database baseline should evaluate");
        assert_number(&baseline, 40.0, &baseline_source);
        let reads = baseline_resolver.reads();
        assert!(
            reads > 0,
            "baseline database evaluation must read the provider"
        );
        drop(baseline);
        assert_eq!(baseline_budget.used(Resource::Memory), 0);
        reads
    };

    let resolver = DatabaseLimitsResolver::standard("East");
    let (budget, _cancellation, execution) = make_execution("ods-formula-database-projected-cache");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("projected database scalar should produce a matrix-shaped result");
    let array = result
        .as_array()
        .expect("matrix IF should retain the condition shape");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 3));
    assert!(matches!(array.cell(0, 0), Some(Value::Number(40.0))));
    assert!(
        matches!(array.cell(0, 1), Some(Value::Number(0.0))),
        "false projected cell must select the scalar alternative"
    );
    assert!(matches!(array.cell(0, 2), Some(Value::Number(40.0))));
    assert_eq!(
        resolver.reads(),
        baseline_reads,
        "projected database branch should be evaluated once and reused"
    );
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn dget_avoids_unrelated_record_fields_when_the_projection_does_not_need_them() {
    let source = format!("=DGET({DATABASE};3;{CRITERIA})");
    let expression = parse(&source);
    let resolver = DatabaseLimitsResolver::standard("West");
    let (_budget, _cancellation, execution) = make_execution("ods-formula-database-field-read");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &ValueLimits::default(),
    )
    .expect("unique numeric DGET should evaluate");
    assert_number(&result, 20.0, &source);
    let coordinates = resolver.read_coordinates();
    assert!(
        !coordinates
            .iter()
            .any(|&(row, column)| row > 0 && column == 0),
        "record names are unrelated to numeric DGET: {coordinates:?}"
    );
    assert!(
        !coordinates
            .iter()
            .any(|&(row, column)| row > 0 && column == 3),
        "unused record fields must not be read: {coordinates:?}"
    );
}

#[test]
fn owned_dget_number_survives_source_and_resolver_drop() {
    let (budget, _cancellation, execution) = make_execution("ods-formula-database-owned-dget");
    let owned = {
        let source = String::from("=DGET([.A1:.D4];3;[.F1:.F2])");
        let expression = parse(&source);
        let resolver = DatabaseLimitsResolver::standard("West");
        let borrowed = evaluate(
            &expression,
            &resolver,
            &execution,
            Mode::Scalar,
            &ValueLimits::default(),
        )
        .expect("unique numeric DGET should evaluate");
        assert!(matches!(borrowed.value(), Value::Number(20.0)));
        let owned = borrowed
            .to_owned(&execution, &ValueLimits::default())
            .expect("numeric DGET should be ownable");
        assert!(matches!(owned.value(), OwnedValueView::Number(20.0)));
        owned
    };

    drop(execution);
    assert!(matches!(owned.value(), OwnedValueView::Number(20.0)));
    assert_eq!(owned.reserved_storage_bytes(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
    drop(owned);
    assert_eq!(budget.used(Resource::Memory), 0);
}
