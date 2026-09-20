//! Independent consumer for the ODF date/time oracle corpus.
//!
//! The expected values are authored by the resolver-free Python oracle in the
//! evidence bundle.  This target replays every retained formula through the
//! Rust value evaluator with a deliberately empty resolver.  The corpus hash
//! recorded here is the reviewed contract identity; the evidence verifier
//! remains responsible for hashing the contract and corpus bytes themselves.

use std::{
    any::Any,
    cell::Cell,
    collections::BTreeSet,
    num::{NonZeroU64, NonZeroUsize},
    path::PathBuf,
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::{
        CalculationTimestamp, EvaluationFailure, EvaluationOptions, ScalarError, UnsupportedKind,
        value::{self, CellRead, Context, Limits, Mode, Position, Resolver, SheetExtent, Value},
    },
    expression::Expression,
};
use serde_json::Value as JsonValue;

const CONTRACT_SHA256: &str = "cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f";

const FUNCTIONS: [&str; 24] = [
    "DATE",
    "DATEDIF",
    "DATEVALUE",
    "DAY",
    "DAYS",
    "DAYS360",
    "EASTERSUNDAY",
    "EDATE",
    "EOMONTH",
    "HOUR",
    "ISOWEEKNUM",
    "MINUTE",
    "MONTH",
    "NETWORKDAYS",
    "NOW",
    "SECOND",
    "TIME",
    "TIMEVALUE",
    "TODAY",
    "WEEKDAY",
    "WEEKNUM",
    "WORKDAY",
    "YEAR",
    "YEARFRAC",
];

#[derive(Debug)]
struct EmptyResolver {
    reads: Cell<usize>,
}

impl EmptyResolver {
    fn new() -> Self {
        Self {
            reads: Cell::new(0),
        }
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }
}

impl Resolver for EmptyResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(SheetExtent::new(16, 16)))
    }

    fn read_cell<'a>(
        &'a self,
        _sheet: &str,
        _row: usize,
        _column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.set(self.reads.get().saturating_add(1));
        Ok(CellRead::Empty)
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(0))
    }

    fn sheet_name_at<'a>(
        &'a self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&'a str>, EvaluationFailure> {
        Ok((index == 0).then_some("Main"))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(1)
    }
}

fn execution() -> ExecutionContext {
    let budget = Budget::root(
        "ods-formula-date-time-independent-oracle",
        CoreLimits::for_profile(Profile::Server),
    );
    let (_cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one in-flight task"),
        NonZeroU64::new(1024 * 1024).expect("one in-flight MiB"),
        0,
    )
    .expect("valid execution limits");
    ExecutionContext::new(budget, token, limits)
}

fn formula_error(value: &str) -> ScalarError {
    match value {
        "#NULL!" => ScalarError::Null,
        "#DIV/0!" => ScalarError::DivisionByZero,
        "#VALUE!" => ScalarError::Value,
        "#REF!" => ScalarError::Reference,
        "#NAME?" => ScalarError::Name,
        "#NUM!" => ScalarError::Number,
        "#N/A" => ScalarError::NotAvailable,
        other => panic!("unsupported date/time-oracle error {other:?}"),
    }
}

fn options(row: &JsonValue) -> EvaluationOptions {
    let Some(timestamp) = row
        .get("context")
        .and_then(|context| context.get("calculation_timestamp"))
        .filter(|timestamp| !timestamp.is_null())
    else {
        return EvaluationOptions::default();
    };
    let serial = timestamp["value"]
        .as_f64()
        .unwrap_or_else(|| panic!("{}: timestamp is not numeric", row["id"]));
    let calculation_timestamp = CalculationTimestamp::from_serial(serial)
        .unwrap_or_else(|error| panic!("{}: invalid timestamp: {error:?}", row["id"]));
    EvaluationOptions::default().with_calculation_timestamp(calculation_timestamp)
}

fn assert_number(expected: &JsonValue, actual: Value<'_>, case: &str) {
    let wanted = expected["value"]
        .as_f64()
        .unwrap_or_else(|| panic!("{case}: numeric expected value is missing"));
    let Value::Number(observed) = actual else {
        panic!("{case}: expected Number({wanted}), got {actual:?}");
    };
    if let Some(exact) = expected.get("exact").and_then(JsonValue::as_str) {
        let exact_value = exact
            .parse::<f64>()
            .unwrap_or_else(|error| panic!("{case}: invalid exact decimal {exact:?}: {error}"));
        assert_close(exact_value, wanted, case, "corpus value/exact mismatch");
    }
    assert_close(observed, wanted, case, "evaluation mismatch");
}

fn assert_close(observed: f64, wanted: f64, case: &str, detail: &str) {
    let scale = observed.abs().max(wanted.abs()).max(1.0);
    assert!(
        (observed - wanted).abs() <= 1e-12 * scale,
        "{case}: {detail}: expected {wanted:?}, observed {observed:?}"
    );
}

fn assert_success(expected: &JsonValue, actual: Value<'_>, case: &str) {
    match expected["kind"]
        .as_str()
        .unwrap_or_else(|| panic!("{case}: expected kind"))
    {
        "number" => assert_number(expected, actual, case),
        "error" => {
            let wanted = formula_error(
                expected["code"]
                    .as_str()
                    .unwrap_or_else(|| panic!("{case}: expected error code")),
            );
            let Value::Error(observed) = actual else {
                panic!("{case}: expected Error({wanted:?}), got {actual:?}");
            };
            assert_eq!(observed, wanted, "{case}: formula error mismatch");
        },
        "unsupported" => panic!("{case}: unsupported result unexpectedly succeeded"),
        other => panic!("{case}: unsupported expected kind {other:?}"),
    }
}

fn panic_message(payload: Box<dyn Any + Send>) -> String {
    if let Some(message) = payload.downcast_ref::<String>() {
        return message.clone();
    }
    if let Some(message) = payload.downcast_ref::<&str>() {
        return (*message).to_owned();
    }
    "non-string panic payload".to_owned()
}

#[test]
fn date_time_oracle_replays_all_contract_vectors() {
    let path: PathBuf = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(
        "../../docs/report/spec-gap-validation-evidence/ods-formula-date-time/oracle-vectors.json",
    );
    let bytes = std::fs::read(&path).unwrap_or_else(|error| {
        panic!(
            "retained date/time oracle is unavailable at {}: {error}",
            path.display()
        )
    });
    let document: JsonValue = serde_json::from_slice(&bytes).expect("date/time oracle JSON");
    assert_eq!(
        document["schema"].as_str(),
        Some("ods-formula-date-time-oracle-v1"),
        "date/time oracle schema"
    );
    assert_eq!(
        document["contract_sha256"].as_str(),
        Some(CONTRACT_SHA256),
        "date/time oracle declared contract identity"
    );
    let rows = document["vectors"]
        .as_array()
        .expect("date/time oracle vectors");
    assert_eq!(
        document["vector_count"].as_u64(),
        Some(rows.len() as u64),
        "date/time oracle vector count"
    );
    assert!(!rows.is_empty(), "date/time oracle vectors are empty");

    let expected_functions: BTreeSet<&str> = FUNCTIONS.into_iter().collect();
    let mut observed_functions = BTreeSet::new();
    let mut observed_ids: BTreeSet<String> = BTreeSet::new();
    let execution = execution();
    let mut failures = Vec::new();

    for row in rows {
        let case = row["id"].as_str().unwrap_or("<missing-id>").to_owned();
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            assert!(
                observed_ids.insert(case.clone()),
                "duplicate date/time oracle id {case:?}"
            );
            let function = row["function"]
                .as_str()
                .unwrap_or_else(|| panic!("{case}: missing function"));
            assert!(
                FUNCTIONS.contains(&function),
                "{case}: function {function:?} is outside the 24-function corpus"
            );
            observed_functions.insert(function);
            let formula = row["formula"]
                .as_str()
                .unwrap_or_else(|| panic!("{case}: missing formula"));
            let expression = Expression::parse(formula)
                .unwrap_or_else(|error| panic!("{case}: {formula} should parse: {error}"));
            let resolver = EmptyResolver::new();
            let context = Context::new(&execution, Position::new("Main", 0, 0))
                .with_options(options(row))
                .with_mode(Mode::Scalar);
            let evaluation = value::evaluate(&expression, &resolver, &context, &Limits::default());
            match row["expected"]["kind"]
                .as_str()
                .unwrap_or_else(|| panic!("{case}: missing expected kind"))
            {
                "unsupported" => {
                    let capability = row["expected"]["capability"]
                        .as_str()
                        .unwrap_or_else(|| panic!("{case}: missing unsupported capability"));
                    assert_eq!(capability, "CalculationClock", "{case}: capability profile");
                    assert!(
                        matches!(
                            evaluation,
                            Err(EvaluationFailure::Unsupported(
                                UnsupportedKind::CalculationClock
                            ))
                        ),
                        "{case}: expected CalculationClock refusal, got {evaluation:?}"
                    );
                },
                "number" | "error" => {
                    let result = evaluation.unwrap_or_else(|error| {
                        panic!("{case}: {formula} evaluation failed: {error:?}")
                    });
                    assert_success(&row["expected"], result.value(), &case);
                },
                other => panic!("{case}: unsupported expected kind {other:?}"),
            }

            if let Some(reads) = row.get("reads") {
                let expected_reads = reads["expected"]
                    .as_u64()
                    .unwrap_or_else(|| panic!("{case}: expected read count"));
                assert_eq!(
                    resolver.reads(),
                    expected_reads as usize,
                    "{case}: resolver reads"
                );
            }
        }));
        if let Err(payload) = result {
            failures.push(format!("{case}: {}", panic_message(payload)));
        }
    }

    if observed_ids.len() != rows.len() {
        failures.push(format!(
            "every retained date/time vector must execute exactly once: {} unique IDs for {} rows",
            observed_ids.len(),
            rows.len()
        ));
    }
    if observed_functions != expected_functions {
        failures.push(format!(
            "oracle must cover all 24 date/time functions: observed {observed_functions:?}"
        ));
    }
    assert!(
        failures.is_empty(),
        "date/time oracle failures ({}):\n{}",
        failures.len(),
        failures.join("\n")
    );
}
