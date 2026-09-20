//! Exact UTF-8 boundary corpus generated independently by byte_oracle.py.
use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Profile};
use litchi_ods::codec::formula::{
    evaluation::{EvaluationContext, EvaluationLimits, ScalarError, ScalarValue, evaluate_scalar},
    expression::Expression,
};
use std::num::{NonZeroU64, NonZeroUsize};

#[test]
fn byte_functions_match_independent_utf8_boundary_oracle() {
    let document: serde_json::Value = serde_json::from_str(include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/byte-goldens.json"
    )).expect("byte oracle");
    let rows = document["observations"].as_array().expect("observations");
    assert_eq!(rows.len(), 1376);
    check_rows(rows);
}

fn check_rows(rows: &[serde_json::Value]) {
    for row in rows {
        let source = row["formula"].as_str().expect("formula");
        let expression = Expression::parse(source).expect("valid fixture expression");
        let budget = Budget::root("byte-text-oracle", Limits::for_profile(Profile::Server));
        let (_cancellation, token) = CancellationSource::pair();
        let execution = ExecutionContext::new(
            budget,
            token,
            ExecutionLimits::new(
                NonZeroUsize::new(1).expect("worker"),
                NonZeroUsize::new(1).expect("task"),
                NonZeroU64::new(1024 * 1024).expect("memory"),
                0,
            )
            .expect("limits"),
        );
        let result = evaluate_scalar(
            &expression,
            &EvaluationContext::new(&execution),
            &EvaluationLimits::default(),
        )
        .unwrap_or_else(|error| panic!("{source}: {error}"));
        let expected = &row["expected"];
        match expected["kind"].as_str().expect("kind") {
            "Text" => match result.value() {
                ScalarValue::Text(actual) => assert_eq!(
                    actual.as_ref(),
                    expected["value"].as_str().expect("text"),
                    "{source}"
                ),
                actual => panic!("{source}: expected text, got {actual:?}"),
            },
            "Number" => match result.value() {
                ScalarValue::Number(actual) => assert_eq!(
                    actual.to_bits(),
                    expected["value"].as_f64().expect("number").to_bits(),
                    "{source}"
                ),
                actual => panic!("{source}: expected number, got {actual:?}"),
            },
            "Error" => assert!(
                matches!(result.value(), ScalarValue::Error(ScalarError::Value)),
                "{source}: {:?}",
                result.value()
            ),
            other => panic!("unexpected oracle kind {other}"),
        }
    }
}

#[test]
fn native_observations_preserve_the_documented_profile_divergences() {
    let rows: Vec<serde_json::Value> = serde_json::from_str(include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/native/native-results.json"
    )).expect("native observations");
    assert_eq!(rows.len(), 19);
    let mut divergences = 0;
    let mut ascii_matches = 0;
    let expected: Vec<_> = rows.iter().map(|row| {
        let matching = row["profile"] == row["native"];
        assert_eq!(row["matches_profile"].as_bool(), Some(matching));
        if !matching { divergences += 1; }
        if row["case"].as_str().expect("case").starts_with("ascii_") {
            assert!(matching);
            ascii_matches += 1;
        }
        let kind = match row["profile"]["type"].as_str().expect("type") {
            "number" => "Number",
            "text" => "Text",
            other => panic!("unexpected profile result {other}"),
        };
        serde_json::json!({
            "formula": row["formula"].as_str().expect("formula").strip_prefix("of:").expect("OpenFormula"),
            "expected": {"kind": kind, "value": row["profile"]["value"]},
        })
    }).collect();
    assert_eq!(ascii_matches, 7);
    assert_eq!(divergences, 10);
    check_rows(&expected);
}
