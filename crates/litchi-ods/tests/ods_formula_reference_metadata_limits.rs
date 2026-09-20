//! Resource and failure-boundary coverage for reference metadata functions.
//!
//! Descriptor operations are intentionally tested with an instrumented
//! resolver.  Geometry, sheet order, and pseudotype refusal must finish before
//! a provider cell read; output-producing ROW/COLUMN calls must charge their
//! result cells and release reservations on every typed failure.

use std::{
    cell::Cell,
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource, SourceVersion,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        CellRead, Context, Limits, Mode, Position, Resolver, SheetExtent, Value, evaluate,
    },
    evaluation::{EvaluationFailure, ScalarError, UnsupportedKind},
    expression::Expression,
};

#[derive(Debug)]
struct MetadataResolver {
    sheets: Vec<String>,
    extent: SheetExtent,
    reads: Cell<usize>,
    metadata_calls: Cell<usize>,
    fail_metadata: Cell<bool>,
    fail_first_read: Cell<bool>,
    fail_first_metadata: Cell<bool>,
    metadata_failure_is_reference: Cell<bool>,
    fail_first_sheet_count: Cell<bool>,
    source_versions: Cell<Option<(SourceVersion, SourceVersion)>>,
    source_version_calls: Cell<usize>,
}

impl MetadataResolver {
    fn standard() -> Self {
        Self {
            sheets: ["Main", "Data", "Hidden", "Archive"]
                .into_iter()
                .map(str::to_owned)
                .collect(),
            extent: SheetExtent::new(512, 64),
            reads: Cell::new(0),
            metadata_calls: Cell::new(0),
            fail_metadata: Cell::new(false),
            fail_first_read: Cell::new(false),
            fail_first_metadata: Cell::new(false),
            metadata_failure_is_reference: Cell::new(false),
            fail_first_sheet_count: Cell::new(false),
            source_versions: Cell::new(None),
            source_version_calls: Cell::new(0),
        }
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn metadata_calls(&self) -> usize {
        self.metadata_calls.get()
    }

    fn fail_metadata(&self) {
        self.fail_metadata.set(true);
    }

    fn set_source_versions(&self, expected: SourceVersion, observed: SourceVersion) {
        self.source_versions.set(Some((expected, observed)));
        self.source_version_calls.set(0);
    }
}

impl Resolver for MetadataResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        Ok(self
            .sheets
            .iter()
            .any(|name| name == sheet)
            .then_some(self.extent))
    }

    fn read_cell<'a>(
        &'a self,
        _sheet: &str,
        _row: usize,
        _column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.set(self.reads.get().saturating_add(1));
        if self.fail_first_read.replace(false) {
            return Err(EvaluationFailure::Unsupported(
                UnsupportedKind::ReferenceOperator,
            ));
        }
        Ok(CellRead::Number(1.0))
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        if self.fail_first_metadata.replace(false) {
            return Err(EvaluationFailure::Unsupported(
                if self.metadata_failure_is_reference.get() {
                    UnsupportedKind::Reference
                } else {
                    UnsupportedKind::ReferenceOperator
                },
            ));
        }
        if self.fail_metadata.get() {
            return Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference));
        }
        Ok(self.sheets.iter().position(|name| name == sheet))
    }

    fn sheet_name_at<'a>(
        &'a self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&'a str>, EvaluationFailure> {
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        if self.fail_metadata.get() {
            return Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference));
        }
        Ok(self.sheets.get(index).map(String::as_str))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        if self.fail_first_sheet_count.replace(false) {
            return Err(EvaluationFailure::Unsupported(
                UnsupportedKind::ReferenceOperator,
            ));
        }
        if self.fail_metadata.get() {
            return Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference));
        }
        Ok(self.sheets.len())
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

fn evaluate_value(
    source: &str,
    resolver: &MetadataResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<Observed, EvaluationFailure> {
    let expression = Expression::parse(source).expect("test formula parses");
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let result = evaluate(&expression, resolver, &context, limits)?;
    Ok(match result.value() {
        Value::Number(value) => Observed::Number(value),
        Value::Error(error) => Observed::Error(error),
        Value::Array(array) => Observed::Array(array.shape().rows(), array.shape().columns()),
        _ => Observed::Other,
    })
}

#[derive(Debug, PartialEq)]
enum Observed {
    Number(f64),
    Error(ScalarError),
    Array(usize, usize),
    Other,
}

#[test]
fn metadata_shape_probe_preserves_provider_reference_operator_failure() {
    for function in ["SHEET", "ROW", "COLUMN"] {
        let resolver = MetadataResolver::standard();
        resolver.fail_first_read.set(true);
        let (budget, _cancellation, execution) = execution("metadata-probe-provider-failure");
        let formula = format!("=IF({{TRUE()}};{function}(IF(SUM([.A1])=1;[.A1];[.A2]));0)");
        let error = evaluate_value(
            &formula,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .expect_err("a provider failure must not become an unavailable shape descriptor");
        assert!(matches!(
            error,
            EvaluationFailure::Unsupported(UnsupportedKind::ReferenceOperator)
        ));
        assert_eq!(
            resolver.reads(),
            1,
            "{function} must not retry the provider"
        );
        assert_eq!(budget.used(Resource::Memory), 0);
    }
}

#[test]
fn metadata_descriptor_probe_preserves_provider_metadata_failure() {
    for function in ["SHEET", "ROW", "COLUMN"] {
        let resolver = MetadataResolver::standard();
        resolver.fail_first_metadata.set(true);
        let (budget, _cancellation, execution) = execution("metadata-probe-lookup-failure");
        let formula = format!("=IF({{TRUE()}};{function}(IF(TRUE();[.A1];[.A2]));0)");
        let error = evaluate_value(
            &formula,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .expect_err("provider metadata failures must escape descriptor discovery");
        assert!(matches!(
            error,
            EvaluationFailure::Unsupported(UnsupportedKind::ReferenceOperator)
        ));
        assert_eq!(resolver.reads(), 0);
        assert_eq!(budget.used(Resource::Memory), 0);
    }
}

#[test]
fn metadata_range_probe_preserves_provider_metadata_failure() {
    let resolver = MetadataResolver::standard();
    resolver.fail_first_sheet_count.set(true);
    let (budget, _cancellation, execution) = execution("metadata-range-probe-failure");
    let error = evaluate_value(
        "=IF({TRUE()};ROW([.A1]:[.A2]);0)",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("range geometry provider failures must escape descriptor discovery");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::ReferenceOperator)
    ));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn computed_metadata_munit_probe_preserves_provider_failure() {
    for formula in [
        "=IF({TRUE()};SHEET(MUNIT([.A1]));0)",
        "=IF({TRUE()};SHEET(MUNIT(N([.A1])));0)",
    ] {
        let resolver = MetadataResolver::standard();
        resolver.fail_first_read.set(true);
        let (budget, _cancellation, execution) = execution("metadata-munit-probe-failure");
        let error = evaluate_value(
            formula,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .expect_err("computed metadata must not retry a failed MUNIT parameter read");
        assert!(matches!(
            error,
            EvaluationFailure::Unsupported(UnsupportedKind::ReferenceOperator)
        ));
        assert_eq!(resolver.reads(), 1);
        assert_eq!(budget.used(Resource::Memory), 0);
    }
}

#[test]
fn computed_metadata_reference_shape_preserves_provider_failure() {
    let resolver = MetadataResolver::standard();
    resolver.fail_first_metadata.set(true);
    resolver.metadata_failure_is_reference.set(true);
    let (budget, _cancellation, execution) = execution("metadata-computed-shape-failure");
    let error = evaluate_value(
        "=IF({TRUE()};SHEET(ABS([.A1:.A2]));0)",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("computed metadata must preserve a provider's reference capability failure");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::Reference)
    ));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn descriptor_geometry_refuses_before_provider_reads_and_releases_budget() {
    let resolver = MetadataResolver::standard();
    let (budget, _cancellation, execution) = execution("reference-metadata-geometry-limit");
    let limits = Limits::default().with_max_reference_cells(3);
    let error = evaluate_value(
        "=COLUMNS([.A1:.D4])",
        &resolver,
        &execution,
        Mode::Scalar,
        &limits,
    )
    .expect_err("the admitted geometry exceeds the reference-cell limit");
    assert!(matches!(
        error,
            EvaluationFailure::ResourceLimit(ref limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0, "geometry refusal precedes cell reads");
    assert_eq!(
        budget.used(Resource::Memory),
        0,
        "failed reservation is refunded"
    );
}

#[test]
fn descriptor_operations_ignore_array_result_limit_but_projected_output_is_bounded() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-output-limit");

    let descriptor = evaluate_value(
        "=COLUMNS([.A1:.D4])",
        &resolver,
        &execution,
        Mode::Scalar,
        &Limits::default().with_max_array_cells(1),
    )
    .expect("a scalar descriptor result does not use the array-cell quota");
    assert_eq!(descriptor, Observed::Number(4.0));
    assert_eq!(resolver.reads(), 0);

    let error = evaluate_value(
        "=IF({TRUE()|TRUE()};COLUMNS([.A1:.D4]);0)",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_array_cells(1),
    )
    .expect_err("the projected two-cell result exceeds the array-cell limit");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn axis_output_and_reference_area_limits_are_checked_before_reads() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-axis-limit");

    for source in ["=ROW([.A1:.A4])", "=COLUMN([.A1:.D1])"] {
        let error = evaluate_value(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default().with_max_array_cells(3),
        )
        .expect_err("axis output must respect max_array_cells");
        assert!(
            matches!(
                &error,
                EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
            ),
            "{source}: {error:?}"
        );
        assert_eq!(resolver.reads(), 0, "{source}");
    }

    let error = evaluate_value(
        "=ROW([.A1:.A4])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("axis generation must charge work");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work
    ));
    assert_eq!(resolver.reads(), 0);

    let error = evaluate_value(
        "=COLUMNS([.A1:.D1])",
        &resolver,
        &execution,
        Mode::Scalar,
        &Limits::default().with_max_reference_areas(0),
    )
    .expect_err("reference-area metadata reservation must be checked");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn metadata_provider_failure_and_cancellation_remain_typed_and_unread() {
    let resolver = MetadataResolver::standard();
    let (_budget, cancellation, execution) = execution("reference-metadata-provider-failure");
    resolver.fail_metadata();
    let error = evaluate_value(
        "=SHEETS([Main.A1:Archive.A1])",
        &resolver,
        &execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("metadata provider failure must bubble");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::Reference)
    ));
    assert_eq!(resolver.reads(), 0);

    cancellation.cancel();
    let error = evaluate_value(
        "=SHEETS([Main.A1:Archive.A1])",
        &resolver,
        &execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("pre-cancelled metadata evaluation must stop");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn source_fence_supersedes_a_successful_descriptor_result() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-source-fence");
    let expected = SourceVersion::new(0x5245_464d, 0);
    let observed = SourceVersion::new(0x5245_464d, 1);
    resolver.set_source_versions(expected, observed);
    let error = evaluate_value(
        "=COLUMNS([.A1:.D4])",
        &resolver,
        &execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("a changed source cannot publish metadata output");
    assert!(matches!(
        error,
        EvaluationFailure::SourceChanged { expected: got_expected, observed: got_observed }
            if got_expected == expected && got_observed == observed
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn source_qualified_sheet_metadata_refusal_precedes_resolver_metadata() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-source-refusal");
    for source in [
        "=SHEET(['file:///book.ods'#.A1])",
        "=SHEETS(['file:///book.ods'#.A1])",
    ] {
        let result = evaluate_value(
            source,
            &resolver,
            &execution,
            Mode::Scalar,
            &Limits::default(),
        )
        .expect("source qualification is a formula-value refusal");
        assert_eq!(result, Observed::Error(ScalarError::Value), "{source}");
        assert_eq!(resolver.reads(), 0, "{source}");
        assert_eq!(resolver.metadata_calls(), 0, "{source}");
    }
}

#[test]
fn wrong_pseudotypes_and_invalid_arity_are_formula_values_without_reads() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-type-limit");
    for source in [
        "=AREAS(1)",
        "=COLUMNS(1)",
        "=ROWS(1)",
        "=SHEETS(1)",
        "=AREAS()",
        "=COLUMNS([.A1];[.B1])",
    ] {
        let result = evaluate_value(
            source,
            &resolver,
            &execution,
            Mode::Scalar,
            &Limits::default(),
        )
        .expect("pseudotype/arity refusal is a formula value");
        assert_eq!(result, Observed::Error(ScalarError::Value), "{source}");
    }
    assert_eq!(resolver.reads(), 0);
}
