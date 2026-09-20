//! Resource, provider, and publication-boundary coverage for the value
//! inspection/conversion family.  The resolver exposes borrowed text and
//! ordered cells so the tests can distinguish streaming inspection from range
//! materialization.

use std::{
    cell::Cell,
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource, SourceVersion,
};
use litchi_ods::codec::formula::{
    evaluation::value::{self, CellRead, Context, Limits, Mode, Position, Resolver, SheetExtent},
    evaluation::{EvaluationContext, EvaluationFailure, EvaluationLimits},
    expression::Expression,
};

#[derive(Debug, Clone)]
enum LimitCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(litchi_ods::codec::formula::evaluation::ScalarError),
    Unsupported,
}

#[derive(Debug)]
struct InspectionLimitResolver {
    rows: usize,
    columns: usize,
    cells: Vec<LimitCell>,
    reads: Cell<usize>,
    cancel_after_read: Option<CancellationSource>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl InspectionLimitResolver {
    fn standard(rows: usize) -> Self {
        let mut resolver = Self {
            rows,
            columns: 4,
            cells: (0..rows * 4).map(|_| LimitCell::Empty).collect(),
            reads: Cell::new(0),
            cancel_after_read: None,
            source_versions: None,
            source_version_calls: Cell::new(0),
        };
        if rows > 0 {
            resolver.set(0, LimitCell::Empty);
        }
        if rows > 1 {
            resolver.set(1, LimitCell::Text(String::new()));
        }
        if rows > 2 {
            resolver.set(2, LimitCell::Number(42.5));
        }
        if rows > 3 {
            resolver.set(3, LimitCell::Logical(true));
        }
        resolver
    }

    fn set(&mut self, row: usize, value: LimitCell) {
        self.cells[row * self.columns] = value;
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn clear_reads(&self) {
        self.reads.set(0);
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

impl Resolver for InspectionLimitResolver {
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
        if let Some(cancellation) = &self.cancel_after_read {
            cancellation.cancel();
        }
        if sheet != "Main" || column != 0 || row >= self.rows {
            return Ok(CellRead::Error(
                litchi_ods::codec::formula::evaluation::ScalarError::Reference,
            ));
        }
        Ok(match &self.cells[row * self.columns] {
            LimitCell::Empty => CellRead::Empty,
            LimitCell::Number(value) => CellRead::Number(*value),
            LimitCell::Logical(value) => CellRead::Logical(*value),
            LimitCell::Text(value) => CellRead::Text(value.as_str()),
            LimitCell::Error(error) => CellRead::Error(*error),
            LimitCell::Unsupported => CellRead::Unsupported,
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

fn evaluate_source(
    source: &str,
    resolver: &InspectionLimitResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<(), EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let _result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(())
}

fn evaluate_scalar_source(
    source: &str,
    execution: &ExecutionContext,
    limits: &EvaluationLimits,
) -> Result<(), EvaluationFailure> {
    let expression = parse(source);
    let _result = litchi_ods::codec::formula::evaluation::evaluate_scalar(
        &expression,
        &EvaluationContext::new(execution),
        limits,
    )?;
    Ok(())
}

fn assert_resource(error: EvaluationFailure, resource: Resource, source: &str) {
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == resource),
        "{source}: expected {resource:?}, got {error:?}"
    );
}

#[test]
fn inspection_shape_and_reference_limits_precede_cell_reads() {
    let resolver = InspectionLimitResolver::standard(4);
    let (budget, _cancellation, execution) = execution("ods-formula-inspection-shapes");
    let error = evaluate_source(
        "=ISNUMBER([.A1:.A4])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("zero reference budget must refuse before a read");
    assert_resource(error, Resource::Objects, "reference cells");
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let error = evaluate_source(
        "=ISNUMBER([.A1:.A4])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_array_cells(1),
    )
    .expect_err("output shape must refuse before a read");
    assert_resource(error, Resource::Objects, "output cells");
    assert_eq!(resolver.reads(), 0);

    evaluate_source(
        "=ISNUMBER([.A1:.A4])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("admitted inspection range");
    assert_eq!(resolver.reads(), 4, "admitted cells stream once each");
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn inspection_scans_charge_work_and_borrowed_text_budget() {
    let rows = 8_193;
    let resolver = InspectionLimitResolver::standard(rows);
    let (budget, _cancellation, work_execution) = execution("ods-formula-inspection-work");
    let error = evaluate_source(
        "=ISNUMBER([.A1:.A8193])",
        &resolver,
        &work_execution,
        Mode::Matrix,
        &Limits::default()
            .with_max_steps(4_096)
            .with_max_reference_cells(rows),
    )
    .expect_err("inspection range work must be bounded");
    assert_resource(error, Resource::Work, "inspection work");
    assert!(resolver.reads() > 0 && resolver.reads() < rows);
    assert_eq!(budget.used(Resource::Memory), 0);

    let mut resolver = InspectionLimitResolver::standard(1);
    resolver.set(0, LimitCell::Text("x".repeat(8_193)));
    let (budget, _cancellation, text_execution) = execution("ods-formula-inspection-text");
    let error = evaluate_source(
        "=ISTEXT([.A1])",
        &resolver,
        &text_execution,
        Mode::Matrix,
        &Limits::default().with_max_text_bytes(8_192),
    )
    .expect_err("borrowed inspection text must honor the text budget");
    assert_resource(error, Resource::Memory, "inspection text bytes");
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn inspection_cancellation_and_source_fences_are_typed() {
    let (cancel_budget, cancellation, cancel_execution) =
        execution("ods-formula-inspection-cancel");
    let resolver = InspectionLimitResolver::standard(1).with_cancel_after_read(&cancellation);
    let error = evaluate_source(
        "=ISNUMBER([.A1])",
        &resolver,
        &cancel_execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("cancellation must fence inspection publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(cancel_budget.used(Resource::Memory), 0);

    let mut resolver = InspectionLimitResolver::standard(1);
    let expected = SourceVersion::new(0x494E_5350, 0);
    let observed = SourceVersion::new(0x494E_5350, 1);
    resolver.set_source_versions(expected, observed);
    let (budget, _cancellation, execution) = execution("ods-formula-inspection-source");
    let error = evaluate_source(
        "=ISNUMBER([.A1])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("source change must fence inspection publication");
    assert!(matches!(
        error,
        EvaluationFailure::SourceChanged {
            expected: got_expected,
            observed: got_observed
        } if got_expected == expected && got_observed == observed
    ));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn typed_provider_failures_escape_inspection_and_iferror() {
    let mut resolver = InspectionLimitResolver::standard(2);
    resolver.set(
        0,
        LimitCell::Error(litchi_ods::codec::formula::evaluation::ScalarError::NotAvailable),
    );
    resolver.set(1, LimitCell::Unsupported);
    let (budget, _cancellation, execution) = execution("ods-formula-inspection-provider");
    let error = evaluate_source(
        "=ISNUMBER([.A1:.A2])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("typed provider failure must supersede a formula error value");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(
            litchi_ods::codec::formula::evaluation::UnsupportedKind::CellValue
        )
    ));
    assert_eq!(resolver.reads(), 2);

    resolver.clear_reads();
    let error = evaluate_source(
        "=IFERROR(ISNUMBER([.A1:.A2]);FALSE())",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("IFERROR must not catch a typed provider failure");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(
            litchi_ods::codec::formula::evaluation::UnsupportedKind::CellValue
        )
    ));
    assert_eq!(resolver.reads(), 2);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn scalar_conversion_text_budget_is_typed_and_refunded() {
    let (budget, _cancellation, execution) = execution("ods-formula-inspection-scalar-budget");
    let error = evaluate_scalar_source(
        r#"=IFERROR(VALUE("123456789");0)"#,
        &execution,
        &EvaluationLimits::default()
            .with_max_text_bytes(3)
            .with_max_storage_bytes(64),
    )
    .expect_err("scalar conversion text must honor the input text budget");
    assert_resource(error, Resource::Memory, "scalar conversion text");
    assert_eq!(budget.used(Resource::Memory), 0);
}
