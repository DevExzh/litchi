//! Focused limits, cancellation, and ownership coverage for matrix formulas.
//!
//! Matrix functions are intentionally exercised through the public value
//! evaluator. Formula errors remain successful values, while evaluator
//! limits, cancellation, and ownership failures remain typed Rust results.

use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource,
};
use litchi_ods::codec::formula::{
    evaluation::EvaluationFailure,
    evaluation::value::{
        CellRead, Context, Evaluated, Limits, Mode, OwnedValueView, Position, Resolver,
        SheetExtent, Value, evaluate as evaluate_value,
    },
    expression::Expression,
};

#[derive(Clone, Debug)]
enum FixtureCell {
    Number(f64),
    Text(String),
}

#[derive(Debug)]
struct FixtureResolver {
    cells: Vec<FixtureCell>,
    columns: usize,
    reads: AtomicUsize,
    cancel_after_read: Option<CancellationSource>,
}

impl FixtureResolver {
    fn number(value: f64) -> Self {
        Self {
            cells: vec![FixtureCell::Number(value); 256],
            columns: 16,
            reads: AtomicUsize::new(0),
            cancel_after_read: None,
        }
    }

    fn text(values: [&str; 4]) -> Self {
        Self {
            cells: values
                .into_iter()
                .map(|value| FixtureCell::Text(value.to_owned()))
                .collect(),
            columns: 2,
            reads: AtomicUsize::new(0),
            cancel_after_read: None,
        }
    }

    fn with_cancel_after_read(mut self, cancellation: &CancellationSource) -> Self {
        self.cancel_after_read = Some(cancellation.clone());
        self
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::Acquire)
    }
}

impl Resolver for FixtureResolver {
    fn sheet_extent(
        &self,
        _sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok(Some(SheetExtent::new(
            self.cells.len() / self.columns,
            self.columns,
        )))
    }

    fn read_cell<'a>(
        &'a self,
        _sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.fetch_add(1, Ordering::AcqRel);
        if let Some(cancellation) = &self.cancel_after_read {
            cancellation.cancel();
        }
        let index = row
            .checked_mul(self.columns)
            .and_then(|index| index.checked_add(column));
        Ok(match index.and_then(|index| self.cells.get(index)) {
            Some(FixtureCell::Number(value)) => CellRead::Number(*value),
            Some(FixtureCell::Text(value)) => CellRead::Text(value.as_str()),
            None => CellRead::Empty,
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

fn new_execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
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

fn evaluate_at<'a>(
    expression: &'a Expression,
    resolver: &'a FixtureResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    evaluate_value(expression, resolver, &context, limits)
}

#[test]
fn matrix_function_families_enforce_max_array_cells_before_allocation() {
    for source in [
        "=MDETERM({1;0|0;1})",
        "=MINVERSE({1;0|0;1})",
        "=MMULT({1;2|3;4};{1;0|0;1})",
        "=MUNIT(2)",
        "=TRANSPOSE({1;2|3;4})",
    ] {
        let (budget, _cancellation, execution) = new_execution("ods-formula-matrix-array-limit");
        let resolver = FixtureResolver::number(0.0);
        let expression = parse(source);
        let error = evaluate_at(
            &expression,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default().with_max_array_cells(3),
        )
        .expect_err("a four-cell matrix must exceed the three-cell limit");
        assert!(
            matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
            "{source:?} returned the wrong limit failure: {error:?}"
        );
        assert_eq!(budget.used(Resource::Memory), 0, "{source:?}");
    }
}

#[test]
fn matrix_domain_errors_remain_formula_values() {
    for source in [
        "=MDETERM({1;2})",
        "=MINVERSE({1;2|2;4})",
        "=MMULT({1;2};{3;4})",
        "=MUNIT(0)",
    ] {
        let (budget, _cancellation, execution) = new_execution("ods-formula-matrix-errors");
        let resolver = FixtureResolver::number(0.0);
        let expression = parse(source);
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should return a formula error value: {error}"));
        assert!(
            matches!(result.value(), Value::Error(_)),
            "{source:?} returned a non-error value: {:?}",
            result.value()
        );
        drop(result);
        assert_eq!(budget.used(Resource::Memory), 0, "{source:?}");
    }
}

#[test]
fn matrix_work_and_storage_limits_are_typed_separately_from_formula_errors() {
    let expression = parse("=MMULT({1;2|3;4};{5;6|7;8})");
    let resolver = FixtureResolver::number(0.0);
    let (budget, _cancellation, execution) = new_execution("ods-formula-matrix-work-limit");
    let limits = Limits::default().with_max_steps(0);
    let error = evaluate_at(&expression, &resolver, &execution, Mode::Matrix, &limits)
        .expect_err("zero scalar work must refuse matrix evaluation");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong matrix work failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let expression = parse("=MINVERSE({1;2|3;4})");
    let resolver = FixtureResolver::number(0.0);
    let (budget, _cancellation, execution) = new_execution("ods-formula-matrix-storage-limit");
    let limits = Limits::default().with_max_storage_bytes(0);
    let error = evaluate_at(&expression, &resolver, &execution, Mode::Matrix, &limits)
        .expect_err("zero scalar storage must refuse matrix evaluation");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong matrix storage failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn multiplication_charges_arithmetic_after_materializing_its_inputs() {
    let expression = parse("=MMULT([.A1:.P16];[.A1:.P16])");
    let resolver = FixtureResolver::number(1.0);
    let (budget, _cancellation, execution) = new_execution("ods-matrix-arithmetic-work");
    // The two input rectangles contain 512 cells in total, while the product
    // needs 4096 multiply-add terms. This limit admits the reads but must
    // refuse the arithmetic rather than returning an uncharged product.
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_steps(2_000),
    )
    .expect_err("matrix arithmetic must consume the finite work budget");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(ref limit) if limit.resource == Resource::Work
    ));
    assert_eq!(
        resolver.reads(),
        512,
        "both inputs were admitted before arithmetic"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn nested_matrix_shape_probes_have_a_finite_depth_limit() {
    let mut source = String::from("1");
    for _ in 0..48 {
        source = format!("MUNIT(MMULT(IF({{TRUE()}};{source};0);{{1}}))");
    }
    let expression = parse(&format!("=IF({{TRUE()}};{source};0)"));
    let resolver = FixtureResolver::number(1.0);
    let (budget, _cancellation, execution) = new_execution("ods-matrix-probe-depth");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("nested probes must refuse before exhausting the Rust stack");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(ref limit) if limit.resource == Resource::Depth
    ));
    assert_eq!(budget.used(Resource::Depth), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn literal_determinant_branch_is_not_recomputed_for_each_selected_cell() {
    let rows = (0..8)
        .map(|row| {
            (0..8)
                .map(|column| if row == column { "2" } else { "1" })
                .collect::<Vec<_>>()
                .join(";")
        })
        .collect::<Vec<_>>()
        .join("|");
    let work_for = |columns: usize| {
        let condition = vec!["TRUE()"; columns].join(";");
        let expression = parse(&format!("=IF({{{condition}}};MDETERM({{{rows}}});0)"));
        let resolver = FixtureResolver::number(0.0);
        let (budget, _cancellation, execution) = new_execution("ods-matrix-branch-reuse");
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap();
        let array = result.as_array().unwrap();
        assert_eq!(array.len(), columns);
        for index in 0..columns {
            assert!(matches!(array.get(index), Some(Value::Number(n)) if (n - 9.0).abs() < 1e-10));
        }
        drop(result);
        assert_eq!(budget.used(Resource::Memory), 0);
        budget.used(Resource::Work)
    };
    let one = work_for(1);
    let sixteen = work_for(16);
    assert!(
        sixteen < one * 2,
        "sixteen outputs should reuse the determinant: {one} versus {sixteen} work units"
    );
}

#[test]
fn pre_cancelled_value_evaluation_stops_before_resolver_access() {
    let (budget, cancellation, execution) = new_execution("ods-formula-matrix-pre-cancel");
    cancellation.cancel();
    let resolver = FixtureResolver::number(7.0);
    let expression = parse("=MDETERM([.A1:.B2])");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("pre-cancelled evaluation must refuse before reading");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Work), 0);
}

#[test]
fn resolver_mid_read_cancellation_is_not_published() {
    let (budget, cancellation, execution) = new_execution("ods-formula-matrix-mid-read-cancel");
    let resolver = FixtureResolver::number(7.0).with_cancel_after_read(&cancellation);
    let expression = parse("=MDETERM([.A1:.B2])");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("cancellation requested by a provider read must fence publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn owned_result_survives_expression_and_resolver_drop_and_releases_budget() {
    let (budget, _cancellation, execution) = new_execution("ods-formula-matrix-owned-result");
    let limits = Limits::default();
    let owned = {
        let source = String::from("=TRANSPOSE([.A1:.B2])");
        let expression = parse(&source);
        let resolver = FixtureResolver::text(["a", "b", "c", "d"]);
        let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
        let borrowed = evaluate_value(&expression, &resolver, &context, &limits)
            .expect("borrowed resolver text matrix should evaluate");
        let array = borrowed
            .as_array()
            .expect("transpose should return an array");
        assert_eq!((array.shape().rows(), array.shape().columns()), (2, 2));
        assert!(matches!(array.cell(0, 0), Some(Value::Text("a"))));
        assert!(matches!(array.cell(0, 1), Some(Value::Text("c"))));
        assert!(matches!(array.cell(1, 0), Some(Value::Text("b"))));
        assert!(matches!(array.cell(1, 1), Some(Value::Text("d"))));
        let owned = borrowed
            .to_owned(&execution, &limits)
            .expect("owned conversion should admit the transposed text array");
        let owned_array = owned
            .as_array()
            .expect("owned result should retain the transposed array");
        assert_eq!(
            (owned_array.shape().rows(), owned_array.shape().columns()),
            (2, 2)
        );
        assert!(matches!(
            owned_array.cell(0, 0),
            Some(OwnedValueView::Text("a"))
        ));
        assert!(matches!(
            owned_array.cell(0, 1),
            Some(OwnedValueView::Text("c"))
        ));
        assert!(matches!(
            owned_array.cell(1, 0),
            Some(OwnedValueView::Text("b"))
        ));
        assert!(matches!(
            owned_array.cell(1, 1),
            Some(OwnedValueView::Text("d"))
        ));
        assert!(owned.reserved_storage_bytes() > 0);
        owned
    };

    drop(execution);
    assert!(budget.used(Resource::Memory) > 0);
    let owned_array = owned
        .as_array()
        .expect("owned array should outlive expression and resolver");
    assert!(matches!(
        owned_array.cell(0, 1),
        Some(OwnedValueView::Text("c"))
    ));
    drop(owned);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn unselected_huge_munit_is_lazy_while_selected_munit_refuses() {
    let limits = Limits::default()
        .with_max_array_cells(64)
        .with_max_storage_bytes(64 * 1024);

    let (budget, _cancellation, execution) = new_execution("ods-formula-matrix-munit-unselected");
    let resolver = FixtureResolver::number(0.0);
    let expression = parse("=IF(FALSE();MUNIT(1024);0)");
    let result = evaluate_at(&expression, &resolver, &execution, Mode::Matrix, &limits)
        .expect("the unselected huge matrix must not be admitted or evaluated");
    assert!(matches!(result.value(), Value::Number(0.0)));
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(resolver.reads(), 0);

    let (budget, _cancellation, execution) = new_execution("ods-formula-matrix-munit-selected");
    let resolver = FixtureResolver::number(0.0);
    let expression = parse("=IF(TRUE();MUNIT(1024);0)");
    let error = evaluate_at(&expression, &resolver, &execution, Mode::Matrix, &limits)
        .expect_err("the selected huge matrix must exceed max_array_cells");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "selected MUNIT returned the wrong failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}
