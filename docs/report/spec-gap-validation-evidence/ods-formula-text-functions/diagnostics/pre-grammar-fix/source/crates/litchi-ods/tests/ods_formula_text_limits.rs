//! Resource and capability-boundary coverage for the OpenFormula text family.
//!
//! These tests keep provider reads observable.  Text mapping must retain
//! reference descriptors rather than materializing a range, charge each
//! borrowed provider value before conversion, release partial output on every
//! failure, and preserve typed provider/cancellation/source failures above
//! formula-level text errors.

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
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, ScalarValue,
        UnsupportedKind, evaluate_scalar,
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
struct TextLimitResolver {
    rows: usize,
    columns: usize,
    cells: Vec<LimitCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(usize, usize)>>,
    cancel_after_read: Option<CancellationSource>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl TextLimitResolver {
    fn standard(rows: usize) -> Self {
        let mut resolver = Self {
            rows,
            columns: 4,
            cells: (0..rows * 4)
                .map(|_| LimitCell::Text("text".to_owned()))
                .collect(),
            reads: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
            cancel_after_read: None,
            source_versions: None,
            source_version_calls: Cell::new(0),
        };
        if rows > 0 {
            resolver.set(0, LimitCell::Text("abcdefg".to_owned()));
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

impl Resolver for TextLimitResolver {
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
        self.read_order.borrow_mut().push((row, column));
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

fn new_execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
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

fn evaluate_scalar_source(
    source: &str,
    execution: &ExecutionContext,
    limits: &EvaluationLimits,
) -> Result<Observed, EvaluationFailure> {
    let expression = parse(source);
    let result = evaluate_scalar(&expression, &EvaluationContext::new(execution), limits)?;
    Ok(match result.value() {
        ScalarValue::Text(value) => Observed::Text(value.to_string()),
        ScalarValue::Error(error) => Observed::Error(*error),
        ScalarValue::Number(_) | ScalarValue::Logical(_) | ScalarValue::Complex(_) => {
            Observed::Other
        },
        _ => Observed::Other,
    })
}

#[derive(Debug, PartialEq)]
enum Observed {
    Text(String),
    Error(ScalarError),
    Array {
        rows: usize,
        columns: usize,
        first_text: Option<String>,
    },
    Other,
}

fn evaluate_source(
    source: &str,
    resolver: &TextLimitResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<Observed, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(match result.value() {
        Value::Text(value) => Observed::Text(value.to_owned()),
        Value::Error(error) => Observed::Error(error),
        Value::Array(array) => Observed::Array {
            rows: array.shape().rows(),
            columns: array.shape().columns(),
            first_text: array.get(0).and_then(|value| match value {
                Value::Text(value) => Some(value.to_owned()),
                _ => None,
            }),
        },
        Value::Empty
        | Value::Number(_)
        | Value::Logical(_)
        | Value::Complex(_)
        | Value::Reference(_)
        | Value::ReferenceList(_) => Observed::Other,
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
fn reference_and_array_shape_limits_refuse_before_text_reads() {
    let resolver = TextLimitResolver::standard(8);
    let (budget, _cancellation, execution) = new_execution("ods-formula-text-shape-limits");
    let error = evaluate_source(
        "=UPPER([.A1:.A8])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("zero reference-cell budget must reject text ranges");
    assert_resource(error, Resource::Objects, "reference cell limit");
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let resolver = TextLimitResolver::standard(8);
    let error = evaluate_source(
        "=UPPER({\"a\"|\"b\"})",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_array_cells(1),
    )
    .expect_err("array shape limit must refuse before mapping");
    assert_resource(error, Resource::Objects, "array cell limit");
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn text_output_expansion_refuses_atomically_and_refunds_storage() {
    let resolver = TextLimitResolver::standard(1);
    let (budget, _cancellation, execution) = new_execution("ods-formula-text-output-limit");
    let error = evaluate_source(
        "=REPT(\"ab\";4)",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_text_bytes(3),
    )
    .expect_err("expanded text must honor the text-byte limit");
    assert_resource(error, Resource::Memory, "text-byte limit");
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(resolver.reads(), 0);

    let error = evaluate_source(
        "=SUBSTITUTE(\"aaaa\";\"a\";\"bbbb\")",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_storage_bytes(2),
    )
    .expect_err("expanded substitution must honor storage");
    assert_resource(error, Resource::Memory, "text storage limit");
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn generated_formatter_output_limits_are_typed_and_refunded() {
    let cases = [
        (r#"=DOLLAR(1.25;2)"#, 4_usize),
        (r#"=FIXED(1234567;2)"#, 8_usize),
        (r#"=TEXT(1e123;"0.00E+00")"#, 8_usize),
    ];
    for (index, (raw_source, max_text_bytes)) in cases.into_iter().enumerate() {
        for wrapped in [false, true] {
            let resolver = TextLimitResolver::standard(1);
            let (budget, _cancellation, execution) =
                new_execution("ods-formula-text-generated-output");
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
                    .with_max_text_bytes(max_text_bytes)
                    .with_max_storage_bytes(64 * 1024),
            )
            .expect_err("generated formatter output must honor max text bytes");
            assert_resource(error, Resource::Memory, &source);
            assert_eq!(resolver.reads(), 0, "case {index}, wrapped {wrapped}");
            assert_eq!(
                budget.used(Resource::Memory),
                0,
                "case {index}, wrapped {wrapped}"
            );
        }
    }

    for wrapped in [false, true] {
        let raw_source = r#"=TEXT(REPT("x";128);"@@@@@@@@")"#;
        let resolver = TextLimitResolver::standard(1);
        let (budget, _cancellation, execution) =
            new_execution("ods-formula-text-generated-output-boundary");
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
                .with_max_steps(1_000)
                .with_max_text_bytes(128)
                .with_max_storage_bytes(64 * 1024),
        )
        .expect_err("formatter output must fail its text cap before work exhaustion");
        assert_resource(error, Resource::Memory, &source);
        assert_eq!(resolver.reads(), 0, "wrapped {wrapped}");
        assert_eq!(budget.used(Resource::Memory), 0, "wrapped {wrapped}");
    }
}

#[test]
fn formatter_output_work_is_bounded_before_large_results_are_built() {
    let cases = [
        (r#"=FIXED(1;10000)"#, 512_u64),
        (r#"=TEXT(1;"0.000000")"#, 32_u64),
    ];
    for (index, (source, max_steps)) in cases.into_iter().enumerate() {
        let resolver = TextLimitResolver::standard(1);
        let (budget, _cancellation, execution) = new_execution("ods-formula-text-format-work");
        let error = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default()
                .with_max_steps(max_steps)
                .with_max_text_bytes(20_000)
                .with_max_storage_bytes(64 * 1024),
        )
        .expect_err("formatter work must stop before output allocation");
        assert_resource(error, Resource::Work, source);
        assert_eq!(resolver.reads(), 0, "case {index}");
        assert_eq!(budget.used(Resource::Memory), 0, "case {index}");
    }
}

#[test]
fn borrowed_provider_text_is_charged_before_mapping_without_output_storage() {
    let resolver = TextLimitResolver::standard(1);
    let (budget, _cancellation, execution) = new_execution("ods-formula-text-borrowed");
    let result = evaluate_source(
        "=T([.A1])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("T should preserve provider text");
    assert_eq!(
        result,
        Observed::Array {
            rows: 1,
            columns: 1,
            first_text: Some("abcdefg".to_owned()),
        }
    );
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);

    let mut resolver = TextLimitResolver::standard(1);
    resolver.set(0, LimitCell::Text("x".repeat(8_193)));
    let (limit_budget, _limit_cancellation, limit_execution) =
        new_execution("ods-formula-text-borrowed-limit");
    let error = evaluate_source(
        "=UPPER([.A1])",
        &resolver,
        &limit_execution,
        Mode::Matrix,
        &Limits::default().with_max_text_bytes(8_192),
    )
    .expect_err("provider text must be bounded before conversion");
    assert_resource(error, Resource::Memory, "borrowed text limit");
    assert_eq!(resolver.reads(), 1);
    assert_eq!(limit_budget.used(Resource::Memory), 0);
}

#[test]
fn borrowed_text_passthrough_charges_work_and_iferror_cannot_catch_it() {
    let sources = [
        r#"=T([.A1])"#,
        r#"=CONCATENATE([.A1];"")"#,
        r#"=CONCATENATE("";[.A1])"#,
        r#"=TEXT([.A1];"@")"#,
    ];
    let large_text = "x".repeat(8_193);
    for (index, source) in sources.into_iter().enumerate() {
        for wrapped in [false, true] {
            let mut resolver = TextLimitResolver::standard(1);
            resolver.set(0, LimitCell::Text(large_text.clone()));
            let (budget, _cancellation, execution) =
                new_execution("ods-formula-text-borrowed-work");
            let source = if wrapped {
                format!(r#"=IFERROR({};"fallback")"#, source.trim_start_matches('='))
            } else {
                source.to_owned()
            };
            let error = evaluate_source(
                &source,
                &resolver,
                &execution,
                Mode::Matrix,
                &Limits::default()
                    .with_max_steps(4_096)
                    .with_max_text_bytes(16_384)
                    .with_max_storage_bytes(16_384),
            )
            .expect_err("borrowed text work must fail before output publication");
            assert_resource(error, Resource::Work, &source);
            assert_eq!(resolver.reads(), 1, "source {index}, wrapped {wrapped}");
            assert_eq!(
                budget.used(Resource::Memory),
                0,
                "source {index}, wrapped {wrapped}"
            );
        }
    }
}

#[test]
fn text_mapping_cancellation_after_a_provider_read_is_typed_and_atomic() {
    let (budget, cancellation, execution) = new_execution("ods-formula-text-cancel");
    let resolver = TextLimitResolver::standard(8).with_cancel_after_read(&cancellation);
    let error = evaluate_source(
        "=UPPER([.A1:.A8])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("provider cancellation must fence text publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn text_source_change_after_streaming_is_not_published() {
    let mut resolver = TextLimitResolver::standard(2);
    let expected = SourceVersion::new(0x5445_5854, 0);
    let observed = SourceVersion::new(0x5445_5854, 1);
    resolver.set_source_versions(expected, observed);
    let (budget, _cancellation, execution) = new_execution("ods-formula-text-source-change");
    let error = evaluate_source(
        "=UPPER([.A1:.A2])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("source change must fence text publication");
    assert!(matches!(
        error,
        EvaluationFailure::SourceChanged {
            expected: got_expected,
            observed: got_observed
        } if got_expected == expected && got_observed == observed
    ));
    assert_eq!(resolver.reads(), 2);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn typed_provider_failures_supersede_retained_formula_errors() {
    let mut resolver = TextLimitResolver::standard(2);
    resolver.set(0, LimitCell::Error(ScalarError::NotAvailable));
    resolver.set(1, LimitCell::Unsupported);
    let (_budget, _cancellation, execution) = new_execution("ods-formula-text-provider-error");
    let error = evaluate_source(
        "=UPPER([.A1:.A2])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("typed provider failure must escape a text mapper");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), 2);
}

#[test]
fn ordinary_text_reference_lists_are_rejected_before_the_first_read() {
    let resolver = TextLimitResolver::standard(8);
    let (_budget, _cancellation, execution) = new_execution("ods-formula-text-list-refusal");
    let result = evaluate_source(
        "=UPPER([.A1]~[.A2])",
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
fn text_work_limit_and_pre_cancel_are_typed_without_residual_storage() {
    let resolver = TextLimitResolver::standard(1);
    let (budget, _cancellation, execution) = new_execution("ods-formula-text-work-limit");
    let error = evaluate_source(
        "=REPT(\"x\";100)",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("zero work must reject text construction");
    assert_resource(error, Resource::Work, "text work limit");
    assert_eq!(budget.used(Resource::Memory), 0);

    let (cancel_budget, cancellation, cancel_execution) =
        new_execution("ods-formula-text-pre-cancel");
    cancellation.cancel();
    let resolver = TextLimitResolver::standard(1);
    let error = evaluate_source(
        "=UPPER([.A1])",
        &resolver,
        &cancel_execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("pre-cancelled evaluation must stop before reading");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(cancel_budget.used(Resource::Memory), 0);
}

#[test]
fn scalar_text_output_limits_are_typed_and_iferror_does_not_catch_them() {
    let (budget, _cancellation, execution) = new_execution("ods-formula-text-scalar-limits");
    let error = evaluate_scalar_source(
        "=REPT(\"x\";100)",
        &execution,
        &EvaluationLimits::default().with_max_text_bytes(3),
    )
    .expect_err("scalar text output must honor max text bytes");
    assert_resource(error, Resource::Memory, "scalar text-byte limit");
    assert_eq!(budget.used(Resource::Memory), 0);

    let error = evaluate_scalar_source(
        "=IFERROR(REPT(\"x\";100);\"fallback\")",
        &execution,
        &EvaluationLimits::default().with_max_text_bytes(3),
    )
    .expect_err("IFERROR must not catch a scalar resource refusal");
    assert_resource(error, Resource::Memory, "scalar IFERROR text-byte limit");
    assert_eq!(budget.used(Resource::Memory), 0);
}
