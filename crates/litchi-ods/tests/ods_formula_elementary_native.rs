//! Selected numeric caches from the upstream LibreOffice function fixtures.
//! These observations are not a native application execution or resave claim.

use std::collections::BTreeSet;
use std::num::{NonZeroU64, NonZeroUsize};

use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Profile};
use litchi_ods::codec::formula::{
    evaluation::{EvaluationLimits, ScalarValue, evaluate_scalar_with_context},
    expression::Expression,
};

#[test]
fn selected_libreoffice_numeric_caches_match_the_scalar_profile() {
    let source = include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-elementary-math/native/cached-results.json"
    );
    let observations: serde_json::Value =
        serde_json::from_str(source).expect("cached observations");
    let rows = observations.as_array().expect("observation array");
    assert_eq!(rows.len(), 56);
    let mut functions = BTreeSet::new();
    let budget = Budget::root(
        "elementary-native-caches",
        Limits::for_profile(Profile::Server),
    );
    let (_cancellation, token) = CancellationSource::pair();
    let execution = ExecutionContext::new(
        budget,
        token,
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)
            .expect("execution limits"),
    );
    for row in rows {
        functions.insert(row["function"].as_str().expect("function"));
        let formula = row["formula"].as_str().expect("formula");
        let expected: f64 = row["cached"]
            .as_str()
            .expect("cached decimal")
            .parse()
            .expect("finite number");
        let expression = Expression::parse(formula).expect("upstream formula parses");
        let result =
            evaluate_scalar_with_context(&expression, &execution, &EvaluationLimits::default())
                .expect("bounded upstream scalar evaluates");
        let ScalarValue::Number(actual) = result.value() else {
            panic!("{formula} produced {:?}; cached {expected}", result.value());
        };
        // Upstream caches are usually serialized to 15 significant decimal
        // digits. Use a relative tolerance, without masking tiny results with
        // a fixed absolute epsilon. Exact cached zeros must remain exact zeros.
        let tolerance = expected.abs() * 1e-13;
        assert!(
            actual.is_finite() && (actual - expected).abs() <= tolerance,
            "{formula}: {actual} versus cached {expected}; source {} row {} column {}",
            row["source"],
            row["row"],
            row["column"]
        );
    }
    assert_eq!(functions.len(), 11);
}
