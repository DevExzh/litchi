//! Exact-rational differential observations over seeded binary64 operands.

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::{
        EvaluationFailure, ScalarError,
        value::{CellRead, Context, Limits, Position, Resolver, SheetExtent, Value, evaluate},
    },
    expression::Expression,
};
use std::num::{NonZeroU64, NonZeroUsize};

struct NoCells;

impl Resolver for NoCells {
    fn sheet_extent(
        &self,
        _: &str,
        _: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok(None)
    }

    fn read_cell<'a>(
        &'a self,
        _: &str,
        _: usize,
        _: usize,
        _: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        panic!("closed numerical oracle must not read cells")
    }

    fn sheet_index(
        &self,
        _: &str,
        _: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        Ok(None)
    }

    fn sheet_name_at(
        &self,
        _: usize,
        _: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        Ok(None)
    }

    fn sheet_count(&self, _: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(0)
    }
}

#[test]
fn seeded_binary64_aggregates_match_exact_rational_observations() {
    let document: serde_json::Value = serde_json::from_str(include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-aggregates/differential-goldens.json"
    ))
    .expect("retained oracle JSON");
    let rows = document["observations"].as_array().expect("oracle rows");
    assert_eq!(rows.len(), 336);
    let budget = Budget::root("aggregate-oracle", CoreLimits::for_profile(Profile::Server));
    let (_cancellation, token) = CancellationSource::pair();
    let execution = ExecutionContext::new(
        budget,
        token,
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)
            .expect("execution limits"),
    );
    let context = Context::new(&execution, Position::new("Main", 0, 0));
    for row in rows {
        let formula = row["formula"].as_str().expect("formula");
        let expression = Expression::parse(formula).expect("closed expression parses");
        let result = evaluate(&expression, &NoCells, &context, &Limits::default())
            .unwrap_or_else(|error| panic!("{formula}: {error}"));
        let Some(expected_hex) = row["expected_bits"].as_str() else {
            assert!(
                matches!(result.value(), Value::Error(ScalarError::Number)),
                "{formula}: expected Number error, got {:?}",
                result.value()
            );
            continue;
        };
        let bits = u64::from_str_radix(expected_hex, 16).expect("binary64 bits");
        let expected = f64::from_bits(bits);
        let Value::Number(actual) = result.value() else {
            panic!("{formula}: expected {expected:e}, got {:?}", result.value());
        };
        assert!(actual.is_finite(), "{formula}: non-finite result");
        if expected == 0.0 {
            assert_eq!(actual, 0.0, "{formula}");
        } else {
            assert_eq!(
                actual.is_sign_negative(),
                expected.is_sign_negative(),
                "{formula}"
            );
            let max_ulps = row["max_ulps"].as_u64().expect("ULP policy");
            assert!(
                actual.to_bits().abs_diff(bits) <= max_ulps,
                "{formula}: {actual:e} versus {expected:e}, allowed {max_ulps} ULP"
            );
        }
    }
}

#[test]
fn targeted_aggregate_cancellation_matches_retained_exact_rational_goldens() {
    let document: serde_json::Value = serde_json::from_str(include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-aggregates/numeric-goldens.json"
    ))
    .expect("retained targeted oracle JSON");
    let rows = document["observations"].as_array().expect("oracle rows");
    assert_eq!(rows.len(), 17);
    let budget = Budget::root(
        "aggregate-targeted-oracle",
        CoreLimits::for_profile(Profile::Server),
    );
    let (_cancellation, token) = CancellationSource::pair();
    let execution = ExecutionContext::new(
        budget,
        token,
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)
            .expect("execution limits"),
    );
    let context = Context::new(&execution, Position::new("Main", 0, 0));
    for row in rows {
        let formula = row["formula"].as_str().expect("formula");
        let expression = Expression::parse(formula).expect("closed expression parses");
        let result = evaluate(&expression, &NoCells, &context, &Limits::default())
            .unwrap_or_else(|error| panic!("{formula}: {error}"));
        let bits = u64::from_str_radix(
            row["rounded_binary64_bits"]
                .as_str()
                .expect("expected bits"),
            16,
        )
        .expect("binary64 bits");
        let expected = f64::from_bits(bits);
        let Value::Number(actual) = result.value() else {
            panic!("{formula}: expected {expected:e}, got {:?}", result.value());
        };
        // Scaled products round each mantissa multiplication. The retained
        // reference instead multiplies exact rationals, so allow eight ULPs.
        let max_ulps = if matches!(row["function"].as_str(), Some("PRODUCT" | "SUMPRODUCT")) {
            8
        } else {
            0
        };
        if expected == 0.0 {
            assert_eq!(actual, 0.0, "{formula}");
        } else {
            assert!(actual.is_finite(), "{formula}: non-finite result");
            assert_eq!(
                actual.is_sign_negative(),
                expected.is_sign_negative(),
                "{formula}"
            );
            assert!(
                actual.to_bits().abs_diff(bits) <= max_ulps,
                "{formula}: {actual:e} versus {expected:e}, allowed {max_ulps} ULP"
            );
        }
    }
}
