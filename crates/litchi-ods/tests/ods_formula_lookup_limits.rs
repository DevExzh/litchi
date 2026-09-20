//! Resource and failure-boundary coverage for the OpenFormula 1.4 lookup
//! family.
//!
//! The resolver records metadata and cell access independently.  Descriptor
//! functions must be able to reject malformed or out-of-bounds requests before
//! a provider read, while a successful reference only incurs a read when a
//! consuming function such as `SUM` asks for cell values.

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
    evaluation::{EvaluationFailure, ScalarError, UnsupportedKind},
    expression::Expression,
};

#[derive(Debug, Clone)]
enum FixtureCell {
    Empty,
    Number(f64),
    Text(String),
    Error(ScalarError),
}

#[derive(Debug)]
struct LimitsResolver {
    sheets: Vec<String>,
    extent: SheetExtent,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    metadata_calls: Cell<usize>,
    read_order: RefCell<Vec<(String, usize, usize)>>,
    fail_read: Cell<Option<UnsupportedKind>>,
    fail_read_after: Cell<Option<(usize, UnsupportedKind)>>,
    fail_metadata: Cell<Option<UnsupportedKind>>,
    fail_missing_metadata: Cell<bool>,
    cancel_metadata: Option<CancellationSource>,
    cancel_after_read: Option<CancellationSource>,
    source_versions: Cell<Option<(SourceVersion, SourceVersion)>>,
    source_version_calls: Cell<usize>,
}

impl LimitsResolver {
    fn standard() -> Self {
        let extent = SheetExtent::new(16, 16);
        let mut resolver = Self {
            sheets: ["Main", "Data", "Hidden", "Archive"]
                .into_iter()
                .map(str::to_owned)
                .collect(),
            extent,
            cells: vec![FixtureCell::Empty; extent.rows() * extent.columns()],
            reads: Cell::new(0),
            metadata_calls: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
            fail_read: Cell::new(None),
            fail_read_after: Cell::new(None),
            fail_metadata: Cell::new(None),
            fail_missing_metadata: Cell::new(false),
            cancel_metadata: None,
            cancel_after_read: None,
            source_versions: Cell::new(None),
            source_version_calls: Cell::new(0),
        };
        for (row, value) in [1.0, 2.0, 2.0, 4.0].into_iter().enumerate() {
            resolver.set(row, 0, FixtureCell::Number(value));
        }
        for (row, value) in [10.0, 20.0, 21.0, 40.0].into_iter().enumerate() {
            resolver.set(row, 2, FixtureCell::Number(value));
        }
        resolver.set(0, 1, FixtureCell::Text("one".to_owned()));
        resolver.set(1, 1, FixtureCell::Text("two-first".to_owned()));
        resolver.set(2, 1, FixtureCell::Text("two-last".to_owned()));
        resolver.set(3, 1, FixtureCell::Text("four".to_owned()));
        resolver.set(0, 3, FixtureCell::Error(ScalarError::NotAvailable));
        resolver.set(1, 3, FixtureCell::Number(1.0));
        resolver
    }

    fn set(&mut self, row: usize, column: usize, value: FixtureCell) {
        let index = row
            .checked_mul(self.extent.columns())
            .and_then(|index| index.checked_add(column))
            .expect("fixture coordinate fits");
        self.cells[index] = value;
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn clear(&self) {
        self.reads.set(0);
        self.metadata_calls.set(0);
        self.read_order.borrow_mut().clear();
    }

    fn fail_reads_with(&self, kind: UnsupportedKind) {
        self.fail_read.set(Some(kind));
    }

    fn fail_reads_after(&self, successful_reads: usize, kind: UnsupportedKind) {
        self.fail_read_after.set(Some((successful_reads, kind)));
    }

    fn fail_metadata_with(&self, kind: UnsupportedKind) {
        self.fail_metadata.set(Some(kind));
    }

    fn fail_missing_metadata(&self) {
        self.fail_missing_metadata.set(true);
    }

    fn cancel_metadata(&mut self, cancellation: &CancellationSource) {
        self.cancel_metadata = Some(cancellation.clone());
    }

    fn cancel_after_read(&mut self, cancellation: &CancellationSource) {
        self.cancel_after_read = Some(cancellation.clone());
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
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        if let Some(kind) = self.fail_metadata.get() {
            return Err(EvaluationFailure::Unsupported(kind));
        }
        if let Some(cancellation) = &self.cancel_metadata {
            cancellation.cancel();
            return Ok(None);
        }
        if sheet == "Missing" && self.fail_missing_metadata.get() {
            return Err(EvaluationFailure::Unsupported(
                UnsupportedKind::ReferenceOperator,
            ));
        }
        Ok(self
            .sheets
            .iter()
            .any(|name| name == sheet)
            .then_some(self.extent))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.set(self.reads.get().saturating_add(1));
        self.read_order
            .borrow_mut()
            .push((sheet.to_owned(), row, column));
        if let Some(kind) = self.fail_read.get() {
            return Err(EvaluationFailure::Unsupported(kind));
        }
        if let Some((successful_reads, kind)) = self.fail_read_after.get() {
            if self.reads.get() > successful_reads {
                return Err(EvaluationFailure::Unsupported(kind));
            }
        }
        if let Some(cancellation) = &self.cancel_after_read {
            cancellation.cancel();
        }
        let Some(sheet_index) = self.sheets.iter().position(|name| name == sheet) else {
            return Ok(CellRead::Error(ScalarError::Reference));
        };
        if row >= self.extent.rows() || column >= self.extent.columns() {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        if sheet_index != 0 {
            return Ok(CellRead::Empty);
        }
        Ok(match &self.cells[row * self.extent.columns() + column] {
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
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        if let Some(kind) = self.fail_metadata.get() {
            return Err(EvaluationFailure::Unsupported(kind));
        }
        Ok(self.sheets.iter().position(|name| name == sheet))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        if let Some(kind) = self.fail_metadata.get() {
            return Err(EvaluationFailure::Unsupported(kind));
        }
        Ok(self.sheets.get(index).map(String::as_str))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        if let Some(kind) = self.fail_metadata.get() {
            return Err(EvaluationFailure::Unsupported(kind));
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

fn parse(source: &str) -> Expression {
    Expression::parse(source).unwrap_or_else(|error| panic!("{source:?} should parse: {error}"))
}

#[derive(Debug, PartialEq)]
enum Observed {
    Number(f64),
    Text(String),
    Error(ScalarError),
    Reference,
    Other,
}

fn observe_source(
    source: &str,
    resolver: &LimitsResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<Observed, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(match result.value() {
        Value::Number(value) => Observed::Number(value),
        Value::Text(value) => Observed::Text(value.to_owned()),
        Value::Error(error) => Observed::Error(error),
        Value::Reference(_) | Value::ReferenceList(_) => Observed::Reference,
        _ => Observed::Other,
    })
}

fn assert_resource(error: EvaluationFailure, resource: Resource, source: &str) {
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == resource),
        "{source}: expected {resource:?}, got {error:?}"
    );
}

#[test]
fn descriptor_results_and_refusals_are_read_free() {
    let resolver = LimitsResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-limit-descriptor");

    for source in [
        "=INDEX([.A1:.C4];2;3)",
        "=OFFSET([.A1:.B2];1;1)",
        r#"=INDIRECT("A1")"#,
    ] {
        resolver.clear();
        assert_eq!(
            observe_source(
                source,
                &resolver,
                &execution,
                Mode::Matrix,
                &Limits::default()
            )
            .expect("descriptor result"),
            Observed::Reference,
            "{source}"
        );
        assert_eq!(resolver.reads(), 0, "{source} must not read cells");
    }

    for source in [
        "=OFFSET([.A1:.B2];100;0)",
        "=OFFSET([.A1];-1;0)",
        "=OFFSET([.A1];0;0;0;1)",
        "=OFFSET([.A1];0;0;1;0)",
        r#"=INDIRECT("not-a-reference")"#,
    ] {
        resolver.clear();
        let result = observe_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .expect("checked formula refusal remains a formula value");
        assert!(matches!(result, Observed::Error(_)), "{source}: {result:?}");
        assert_eq!(resolver.reads(), 0, "{source} must refuse before reads");
    }
}

#[test]
fn reference_cell_and_array_limits_are_charged_before_reads() {
    let resolver = LimitsResolver::standard();
    let (budget, _cancellation, reference_execution) = execution("lookup-limit-reference-cells");
    let error = observe_source(
        "=SUM(INDEX([.A1:.C4];2;3))",
        &resolver,
        &reference_execution,
        Mode::Scalar,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("reference-cell limit must reject before the consumer reads");
    assert_resource(error, Resource::Objects, "INDEX reference-cell limit");
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (array_budget, _cancellation, array_execution) = execution("lookup-limit-array-cells");
    let error = observe_source(
        "=INDEX({1;2|3;4};;)",
        &resolver,
        &array_execution,
        Mode::Matrix,
        &Limits::default().with_max_array_cells(3),
    )
    .expect_err("array result limit must reject before materialization");
    assert_resource(error, Resource::Objects, "INDEX array-cell limit");
    assert_eq!(resolver.reads(), 0);
    assert_eq!(array_budget.used(Resource::Memory), 0);
}

#[test]
fn address_text_and_lookup_work_limits_release_memory() {
    let resolver = LimitsResolver::standard();
    let (text_budget, _cancellation, text_execution) = execution("lookup-limit-address-text");
    let error = observe_source(
        "=ADDRESS(123456;123456;1;TRUE())",
        &resolver,
        &text_execution,
        Mode::Scalar,
        &Limits::default().with_max_text_bytes(1),
    )
    .expect_err("ADDRESS text output must respect the text budget");
    assert_resource(error, Resource::Memory, "ADDRESS text limit");
    assert_eq!(resolver.reads(), 0);
    assert_eq!(text_budget.used(Resource::Memory), 0);

    let (work_budget, _cancellation, work_execution) = execution("lookup-limit-work");
    let error = observe_source(
        "=MATCH(3;[.A1:.A4];1)",
        &resolver,
        &work_execution,
        Mode::Scalar,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("lookup scan must charge work before provider reads");
    assert_resource(error, Resource::Work, "MATCH work limit");
    assert_eq!(resolver.reads(), 0);
    assert_eq!(work_budget.used(Resource::Memory), 0);

    let (parse_budget, _cancellation, parse_execution) = execution("lookup-limit-parser");
    let error = observe_source(
        r#"=INDIRECT("R[123456]C[123456]";FALSE())"#,
        &resolver,
        &parse_execution,
        Mode::Matrix,
        &Limits::default().with_max_text_bytes(1),
    )
    .expect_err("INDIRECT normalization must honor its text budget before publication");
    assert_resource(error, Resource::Memory, "INDIRECT text limit");
    assert_eq!(parse_budget.used(Resource::Memory), 0);

    let (storage_budget, _cancellation, storage_execution) =
        execution("lookup-limit-parser-storage");
    let error = observe_source(
        r#"=INDIRECT("A1")"#,
        &resolver,
        &storage_execution,
        Mode::Matrix,
        &Limits::default().with_max_storage_bytes(0),
    )
    .expect_err("INDIRECT parser metadata must reserve storage atomically");
    assert_resource(error, Resource::Memory, "INDIRECT parser storage");
    assert_eq!(storage_budget.used(Resource::Memory), 0);

    let (work_budget, _cancellation, work_execution) = execution("lookup-limit-parser-work");
    let error = observe_source(
        r#"=INDIRECT("R[123456]C[123456]";FALSE())"#,
        &resolver,
        &work_execution,
        Mode::Matrix,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("INDIRECT R1C1 normalization must charge work");
    assert_resource(error, Resource::Work, "INDIRECT parser work");
    assert_eq!(work_budget.used(Resource::Memory), 0);

    let (cancel_budget, cancellation, cancel_execution) = execution("lookup-limit-parser-cancel");
    cancellation.cancel();
    let error = observe_source(
        r#"=INDIRECT("R[123456]C[123456]";FALSE())"#,
        &resolver,
        &cancel_execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("INDIRECT normalization must honor cancellation");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(cancel_budget.used(Resource::Memory), 0);
}

#[test]
fn projected_lookup_probe_limits_drop_partial_cache_state() {
    let formula = "=SUM(IF({TRUE()|TRUE()};INDIRECT([.H1]);0))";

    // The 2x1 condition fits, but the INDIRECT probe cache cannot retain the
    // widened 2x2 demand under this array bound.  Its partial entries and
    // reservations must be released on the typed failure.
    let mut resolver = LimitsResolver::standard();
    resolver.set(0, 7, FixtureCell::Text("A1:B2".to_owned()));
    let (array_budget, _cancellation, array_execution) =
        execution("lookup-limit-probe-array-cache");
    let error = observe_source(
        formula,
        &resolver,
        &array_execution,
        Mode::Matrix,
        &Limits::default().with_max_array_cells(2),
    )
    .expect_err("the widened lookup probe must honor the array-cell bound");
    assert_resource(error, Resource::Objects, "lookup probe array cache");
    assert_eq!(array_budget.used(Resource::Memory), 0);

    // A work refusal during shape probing must also drop any cache entries
    // admitted before the next bounded step.
    let mut resolver = LimitsResolver::standard();
    resolver.set(0, 7, FixtureCell::Text("A1:B2".to_owned()));
    let (work_budget, _cancellation, work_execution) = execution("lookup-limit-probe-work-cache");
    let error = observe_source(
        formula,
        &resolver,
        &work_execution,
        Mode::Matrix,
        &Limits::default().with_max_steps(32),
    )
    .expect_err("lookup probe work must remain bounded");
    assert_resource(error, Resource::Work, "lookup probe work");
    assert_eq!(work_budget.used(Resource::Memory), 0);

    // Cancellation after the first selector read fences the probe before a
    // retained descriptor can be published, and must release its storage.
    let mut resolver = LimitsResolver::standard();
    resolver.set(0, 7, FixtureCell::Text("A1:B2".to_owned()));
    let (cancel_budget, cancellation, cancel_execution) =
        execution("lookup-limit-probe-cancel-cache");
    resolver.cancel_after_read(&cancellation);
    let error = observe_source(
        formula,
        &resolver,
        &cancel_execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("lookup probe cancellation must supersede retained shape state");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(cancel_budget.used(Resource::Memory), 0);

    // The cumulative provider-read limit remains typed even when the lookup
    // probe has already retained selector and descriptor state. Four selector
    // coordinates plus four target cells reach the seventh-read boundary.
    let mut resolver = LimitsResolver::standard();
    resolver.set(0, 7, FixtureCell::Text("A1:B2".to_owned()));
    let (read_budget, _cancellation, read_execution) =
        execution("lookup-limit-probe-reference-reads");
    let error = observe_source(
        formula,
        &resolver,
        &read_execution,
        Mode::Matrix,
        &Limits::default().with_max_reference_cells(7),
    )
    .expect_err("lookup probe reads must honor the cumulative reference limit");
    assert_resource(error, Resource::Objects, "lookup probe reference reads");
    assert_eq!(resolver.reads(), 7);
    assert_eq!(read_budget.used(Resource::Memory), 0);
}

#[test]
fn cancellation_and_source_fences_supersede_lookup_values() {
    let resolver = LimitsResolver::standard();
    let (cancel_budget, cancellation, cancel_execution) = execution("lookup-limit-cancel-before");
    cancellation.cancel();
    let error = observe_source(
        r#"=SUM(INDIRECT("A1"))"#,
        &resolver,
        &cancel_execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("pre-cancelled lookup must stop");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(cancel_budget.used(Resource::Memory), 0);

    let (after_budget, after_cancellation, after_execution) =
        execution("lookup-limit-cancel-after");
    let mut resolver = LimitsResolver::standard();
    resolver.cancel_after_read(&after_cancellation);
    let error = observe_source(
        r#"=SUM(INDIRECT("A1"))"#,
        &resolver,
        &after_execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("cancellation after a provider read must stop publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(after_budget.used(Resource::Memory), 0);

    let (lookup_budget, lookup_cancellation, lookup_execution) =
        execution("lookup-limit-cancel-search");
    let mut resolver = LimitsResolver::standard();
    resolver.cancel_after_read(&lookup_cancellation);
    let error = observe_source(
        "=LOOKUP(1;[.A1:.A4];[.A1])",
        &resolver,
        &lookup_execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("lookup search cancellation must fence extension publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(lookup_budget.used(Resource::Memory), 0);

    let expected = SourceVersion::new(0x4c_4f_4f_4b, 0);
    let observed = SourceVersion::new(0x4c_4f_4f_4b, 1);
    let (source_budget, _cancellation, source_execution) = execution("lookup-limit-source");
    let resolver = LimitsResolver::standard();
    resolver.set_source_versions(expected, observed);
    let error = observe_source(
        r#"=SUM(INDIRECT("A1"))"#,
        &resolver,
        &source_execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("source change must fence lookup publication");
    assert!(matches!(
        error,
        EvaluationFailure::SourceChanged {
            expected: got_expected,
            observed: got_observed
        } if got_expected == expected && got_observed == observed
    ));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(source_budget.used(Resource::Memory), 0);

    let (metadata_budget, metadata_cancellation, metadata_execution) =
        execution("lookup-limit-cancel-metadata");
    let mut resolver = LimitsResolver::standard();
    resolver.cancel_metadata(&metadata_cancellation);
    let error = observe_source(
        "=OFFSET([.A1];0;0)",
        &resolver,
        &metadata_execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("metadata cancellation must survive an extent provider returning None");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(metadata_budget.used(Resource::Memory), 0);
}

#[test]
fn typed_provider_failures_escape_formula_error_handlers_without_retry() {
    let resolver = LimitsResolver::standard();
    resolver.fail_reads_with(UnsupportedKind::CellValue);
    let (_budget, _cancellation, read_execution) = execution("lookup-limit-read-provider");
    let error = observe_source(
        r#"=IFERROR(SUM(INDIRECT("A1"));99)"#,
        &resolver,
        &read_execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("provider cell failure must not become an IFERROR fallback");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), 1);

    let resolver = LimitsResolver::standard();
    resolver.fail_metadata_with(UnsupportedKind::ReferenceOperator);
    let (_budget, _cancellation, metadata_execution) = execution("lookup-limit-metadata-provider");
    let error = observe_source(
        "=INDEX([.A1:.C4];2;3)",
        &resolver,
        &metadata_execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("metadata failure must remain typed");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::ReferenceOperator)
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn typed_failure_after_formula_error_supersedes_retained_search_error() {
    let resolver = LimitsResolver::standard();
    resolver.fail_reads_after(1, UnsupportedKind::CellValue);
    let (_budget, _cancellation, execution) = execution("lookup-limit-error-precedence");
    let error = observe_source(
        "=MATCH(1;[.D1:.D2];0)",
        &resolver,
        &execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("a later typed provider failure must supersede a retained formula error");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(
        resolver.reads(),
        2,
        "the search continues after the first formula error"
    );
}

#[test]
fn unselected_choose_and_if_references_do_not_probe_missing_metadata() {
    let resolver = LimitsResolver::standard();
    resolver.fail_missing_metadata();
    let (_budget, _cancellation, execution) = execution("lookup-limit-lazy");

    for source in [
        "=IF(FALSE();SUM([Missing.A1]);0)",
        "=CHOOSE(1;0;SUM([Missing.A1]))",
    ] {
        assert_eq!(
            observe_source(
                source,
                &resolver,
                &execution,
                Mode::Scalar,
                &Limits::default()
            )
            .expect("unselected branch"),
            Observed::Number(0.0),
            "{source}"
        );
        assert_eq!(resolver.reads(), 0, "{source}");
        resolver.clear();
    }
}

#[test]
fn owned_derived_reference_storage_limit_is_atomic() {
    let resolver = LimitsResolver::standard();
    let (budget, _cancellation, execution) = execution("lookup-limit-owned-storage");
    let expression = parse("=OFFSET([.A1:.B2];1;1)");
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
        .expect("OFFSET descriptor");
    assert!(matches!(result.value(), Value::Reference(_)));
    assert_eq!(resolver.reads(), 0);

    let memory_before_to_owned = budget.used(Resource::Memory);
    let error = result
        .to_owned(&execution, &Limits::default().with_max_storage_bytes(0))
        .expect_err("zero owned-storage budget must reject retained metadata");
    assert_resource(error, Resource::Memory, "owned derived reference");
    assert_eq!(budget.used(Resource::Memory), memory_before_to_owned);
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);
}
