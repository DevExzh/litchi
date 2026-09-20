//! Resource and source-safety coverage for the UTF-8 byte-position text
//! functions.  References stay observable so byte scans cannot silently
//! materialize or clone provider text, and typed failures remain above formula
//! errors and IFERROR.

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
        self, CellRead, Context, Limits, Mode, Position, Resolver, SheetExtent, Value,
    },
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, ScalarValue,
        evaluate_scalar,
    },
    expression::Expression,
};

#[derive(Debug, Clone)]
enum LimitCell {
    Text(String),
    Error(ScalarError),
    Unsupported,
}

#[derive(Debug)]
struct ByteLimitResolver {
    rows: usize,
    columns: usize,
    cells: Vec<LimitCell>,
    reads: Cell<usize>,
    cancel_after_read: Option<CancellationSource>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl ByteLimitResolver {
    fn standard(rows: usize) -> Self {
        let mut resolver = Self {
            rows,
            columns: 4,
            cells: (0..rows * 4)
                .map(|_| LimitCell::Text("text".to_owned()))
                .collect(),
            reads: Cell::new(0),
            cancel_after_read: None,
            source_versions: None,
            source_version_calls: Cell::new(0),
        };
        if rows > 0 {
            resolver.set(0, LimitCell::Text("Aé界🙂".to_owned()));
        }
        resolver
    }

    fn set(&mut self, row: usize, value: LimitCell) {
        self.cells[row * self.columns] = value;
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

impl Resolver for ByteLimitResolver {
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
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        Ok(match &self.cells[row * self.columns] {
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

#[derive(Debug, PartialEq)]
enum Observed {
    Text(String),
    Number(f64),
    Error(ScalarError),
    Array {
        rows: usize,
        columns: usize,
        first_number: Option<f64>,
    },
    Other,
}

fn evaluate_source(
    source: &str,
    resolver: &ByteLimitResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<Observed, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(match result.value() {
        Value::Text(value) => Observed::Text(value.to_owned()),
        Value::Number(value) => Observed::Number(value),
        Value::Error(error) => Observed::Error(error),
        Value::Array(array) => Observed::Array {
            rows: array.shape().rows(),
            columns: array.shape().columns(),
            first_number: array.get(0).and_then(|value| match value {
                Value::Number(value) => Some(value),
                _ => None,
            }),
        },
        Value::Empty
        | Value::Logical(_)
        | Value::Complex(_)
        | Value::Reference(_)
        | Value::ReferenceList(_) => Observed::Other,
        _ => Observed::Other,
    })
}

fn evaluate_scalar_source(
    source: &str,
    execution: &ExecutionContext,
    limits: &EvaluationLimits,
) -> Result<Observed, EvaluationFailure> {
    let expression = parse(source);
    let result = evaluate_scalar(&expression, &EvaluationContext::new(execution), limits)?;
    Ok(match result.value() {
        ScalarValue::Text(value) => Observed::Text(value.to_string()),
        ScalarValue::Number(value) => Observed::Number(*value),
        ScalarValue::Error(error) => Observed::Error(*error),
        ScalarValue::Logical(_) | ScalarValue::Complex(_) => Observed::Other,
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
fn byte_reference_and_list_limits_refuse_before_reads() {
    let resolver = ByteLimitResolver::standard(4);
    let (budget, _cancellation, execution) = execution("ods-formula-byte-text-shapes");
    let error = evaluate_source(
        "=LENB([.A1])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("zero reference budget must refuse before reading");
    assert_resource(error, Resource::Objects, "reference cells");
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let result = evaluate_source(
        "=LENB([.A1]~[.A2])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("ReferenceList refusal is a formula value");
    assert_eq!(result, Observed::Error(ScalarError::Value));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn byte_scans_charge_work_but_successful_len_has_no_output_storage() {
    let mut resolver = ByteLimitResolver::standard(1);
    resolver.set(0, LimitCell::Text("x".repeat(8_193)));
    let (budget, _cancellation, execution) = execution("ods-formula-byte-text-scan");
    let result = evaluate_source(
        "=LENB([.A1])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_steps(20_000),
    )
    .expect("bounded UTF-8 byte count");
    assert_eq!(
        result,
        Observed::Array {
            rows: 1,
            columns: 1,
            first_number: Some(8_193.0),
        }
    );
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);

    let error = evaluate_source(
        "=LENB([.A1])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_steps(4_096),
    )
    .expect_err("byte scan work must be bounded");
    assert_resource(error, Resource::Work, "byte scan work");
    assert_eq!(resolver.reads(), 2);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn byte_output_storage_failure_is_typed_refunded_and_not_catchable() {
    let resolver = ByteLimitResolver::standard(1);
    let (budget, _cancellation, execution) = execution("ods-formula-byte-text-output");
    for raw_source in [r#"=LEFTB("Aé";3)"#, r#"=REPLACEB("Aé";1;1;"long")"#] {
        for wrapped in [false, true] {
            let source = if wrapped {
                format!(
                    r#"=IFERROR({};"fallback")"#,
                    raw_source.trim_start_matches('=')
                )
            } else {
                raw_source.to_owned()
            };
            let error = evaluate_source(
                &source,
                &resolver,
                &execution,
                Mode::Matrix,
                &Limits::default()
                    .with_max_text_bytes(1)
                    .with_max_storage_bytes(64),
            )
            .expect_err("byte output must honor text limit");
            assert_resource(error, Resource::Memory, &source);
            assert_eq!(budget.used(Resource::Memory), 0);
        }
    }
}

#[test]
fn cancellation_and_source_fences_survive_byte_reference_scans() {
    let (cancel_budget, cancellation, cancel_execution) = execution("ods-formula-byte-text-cancel");
    let resolver = ByteLimitResolver::standard(1).with_cancel_after_read(&cancellation);
    let error = evaluate_source(
        "=LENB([.A1])",
        &resolver,
        &cancel_execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("cancellation must fence byte publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(cancel_budget.used(Resource::Memory), 0);

    let mut resolver = ByteLimitResolver::standard(1);
    let expected = SourceVersion::new(0x42595445, 0);
    let observed = SourceVersion::new(0x42595445, 1);
    resolver.set_source_versions(expected, observed);
    let (budget, _cancellation, execution) = execution("ods-formula-byte-text-source");
    let error = evaluate_source(
        "=LENB([.A1])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("source change must fence byte publication");
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
fn typed_provider_failures_supersede_formula_errors_and_scalar_limits_escape_iferror() {
    let mut resolver = ByteLimitResolver::standard(2);
    resolver.set(0, LimitCell::Error(ScalarError::NotAvailable));
    resolver.set(1, LimitCell::Unsupported);
    let (_budget, _cancellation, provider_execution) = execution("ods-formula-byte-text-provider");
    let error = evaluate_source(
        "=LENB([.A1:.A2])",
        &resolver,
        &provider_execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("typed provider failure must supersede retained formula error");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(
            litchi_ods::codec::formula::evaluation::UnsupportedKind::CellValue
        )
    ));
    assert_eq!(resolver.reads(), 2);

    let (budget, _cancellation, execution) = execution("ods-formula-byte-text-scalar-limit");
    let error = evaluate_scalar_source(
        r#"=IFERROR(LEFTB("Aé";3);"fallback")"#,
        &execution,
        &EvaluationLimits::default().with_max_text_bytes(1),
    )
    .expect_err("typed scalar resource failure must escape IFERROR");
    assert_resource(error, Resource::Memory, "scalar byte output");
    assert_eq!(budget.used(Resource::Memory), 0);
}
