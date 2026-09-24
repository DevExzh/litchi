//! Scalar admission for OpenFormula sequence and forced-array aggregates.
//!
//! `SUM`, `PRODUCT`, and `SUMSQ` consume scalar NumberSequence arguments in
//! this resolver-free profile. The four forced-array signatures also accept a
//! scalar as the normative 1x1 array; actual arrays and references are sent to
//! the value VM, where shape and NumberSequence rules are available. This
//! keeps a scalar call useful without turning a matrix into an accidental
//! broadcast reduction.

use super::numerics::{
    NumericAggregate, NumericOperation, ProductSumError, ProductTerm, ScaledProductSum, WideSum,
};
use super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, UnsupportedKind,
    WorkingValue,
};
use litchi_core::Resource;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Function {
    Sum,
    Product,
    SumSquares,
    SumProduct,
    SumX2MY2,
    SumX2PY2,
    SumXMY2,
}

fn function(name: &str) -> Option<Function> {
    Some(if name.eq_ignore_ascii_case("SUM") {
        Function::Sum
    } else if name.eq_ignore_ascii_case("PRODUCT") {
        Function::Product
    } else if name.eq_ignore_ascii_case("SUMSQ") {
        Function::SumSquares
    } else if name.eq_ignore_ascii_case("SUMPRODUCT") {
        Function::SumProduct
    } else if name.eq_ignore_ascii_case("SUMX2MY2") {
        Function::SumX2MY2
    } else if name.eq_ignore_ascii_case("SUMX2PY2") {
        Function::SumX2PY2
    } else if name.eq_ignore_ascii_case("SUMXMY2") {
        Function::SumXMY2
    } else {
        return None;
    })
}

pub(super) fn is_aggregate_function(name: &str) -> bool {
    function(name).is_some()
}

/// Apply one scalar aggregate after eager argument evaluation.
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let Some(function) = function(name) else {
        return Err(EvaluationFailure::Unsupported(UnsupportedKind::Function));
    };
    match function {
        Function::Sum | Function::Product | Function::SumSquares => {
            apply_number_sequence(evaluator, node, function)
        },
        Function::SumProduct | Function::SumX2MY2 | Function::SumX2PY2 | Function::SumXMY2 => {
            apply_forced_array_scalar(evaluator, node, function)
        },
    }
}

fn apply_number_sequence<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    function: Function,
) -> EvaluationResult<()> {
    let count = node.child_count();
    if count == 0 {
        return if function == Function::Product {
            evaluator.finish_invalid_arity(node)
        } else {
            evaluator.push_value(WorkingValue::Number(0.0))
        };
    }
    reverse_value_tail(evaluator, count)?;

    let mut numeric = match function {
        Function::Sum => Some(NumericAggregate::new(NumericOperation::Sum)),
        Function::Product => Some(NumericAggregate::new(NumericOperation::Product)),
        Function::SumSquares => None,
        _ => {
            return Err(EvaluationFailure::InvalidExpression(
                "non-sequence aggregate reached numeric dispatch",
            ));
        },
    };
    let mut squares = (function == Function::SumSquares).then(WideSum::default);
    let mut formula_error = None;
    let mut generated_error = None;

    for _ in 0..count {
        evaluator.charge_work(1)?;
        let value = evaluator.pop_value()?;
        match value {
            WorkingValue::Error(error) => remember(&mut formula_error, error),
            value => match super::to_number(value, evaluator)? {
                Ok(value) => {
                    if let Some(aggregate) = numeric.as_mut() {
                        if let Err(error) = aggregate.push_number(value) {
                            remember(&mut generated_error, error);
                        }
                    } else if let Some(squares) = squares.as_mut()
                        && let Err(error) = squares.push_square(value)
                    {
                        remember(&mut generated_error, error);
                    }
                },
                Err(error) => remember(&mut generated_error, error),
            },
        }
    }

    let value = if let Some(error) = formula_error.or(generated_error) {
        WorkingValue::Error(error)
    } else if let Some(aggregate) = numeric {
        evaluator.charge_work(1)?;
        match aggregate.result() {
            Ok(result) => {
                WorkingValue::Number(result.unwrap_or(if function == Function::Product {
                    1.0
                } else {
                    0.0
                }))
            },
            Err(error) => WorkingValue::Error(error),
        }
    } else if let Some(squares) = squares {
        evaluator.charge_work(1)?;
        match squares.result() {
            Ok(result) => WorkingValue::Number(result.unwrap_or(0.0)),
            Err(error) => WorkingValue::Error(error),
        }
    } else {
        return Err(EvaluationFailure::InvalidExpression(
            "sequence aggregate accumulator is missing",
        ));
    };
    evaluator.push_value(value)
}

/// Evaluate the scalar 1x1 instance of a `ForceArray` signature. An actual
/// array or reference never reaches this function because the scalar visitor
/// reports that capability at the value node; the resolver-aware value VM
/// owns all multi-cell shape and source-order traversal.
fn apply_forced_array_scalar<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    function: Function,
) -> EvaluationResult<()> {
    let count = node.child_count();
    let valid_arity = match function {
        Function::SumProduct => count >= 1,
        Function::SumX2MY2 | Function::SumX2PY2 | Function::SumXMY2 => count == 2,
        _ => false,
    };
    if !valid_arity {
        return evaluator.finish_invalid_arity(node);
    }
    reverse_value_tail(evaluator, count)?;

    let mut formula_error = None;
    let mut generated_error = None;
    let mut span_failure = None;

    if function == Function::SumProduct {
        match count {
            1 => {
                let mut sum = NumericAggregate::new(NumericOperation::Sum);
                let value = pop_number(evaluator, &mut formula_error, &mut generated_error)?;
                if let Some(value) = value {
                    if let Err(error) = sum.push_number(value) {
                        remember(&mut generated_error, error);
                    }
                }
                return finish_numeric_scalar(evaluator, sum, formula_error, generated_error);
            },
            2 => {
                let left = pop_number(evaluator, &mut formula_error, &mut generated_error)?;
                let right = pop_number(evaluator, &mut formula_error, &mut generated_error)?;
                let mut sum = WideSum::default();
                if formula_error.is_none()
                    && generated_error.is_none()
                    && let (Some(left), Some(right)) = (left, right)
                    && let Err(error) = sum.push_product(left, right)
                {
                    remember(&mut generated_error, error);
                }
                return finish_wide_scalar(evaluator, sum, formula_error, generated_error);
            },
            _ => {
                let mut term = ProductTerm::default();
                let mut term_valid = true;
                for _ in 0..count {
                    let value = pop_number(evaluator, &mut formula_error, &mut generated_error)?;
                    if let Some(value) = value {
                        if term.push_factor(value).is_err() {
                            term_valid = false;
                            remember(&mut generated_error, ScalarError::Number);
                        }
                    } else {
                        term_valid = false;
                    }
                }
                if term_valid && formula_error.is_none() && generated_error.is_none() {
                    let mut sum = ScaledProductSum::default();
                    note_product_result(
                        sum.push_term(term),
                        &mut generated_error,
                        &mut span_failure,
                    );
                    if formula_error.is_none() && generated_error.is_none() {
                        match sum.result() {
                            Ok(Some(value)) => {
                                return evaluator.push_value(WorkingValue::Number(value));
                            },
                            Ok(None) => return evaluator.push_value(WorkingValue::Number(0.0)),
                            Err(error) => {
                                note_product_result(
                                    Err(error),
                                    &mut generated_error,
                                    &mut span_failure,
                                );
                            },
                        }
                    }
                }
            },
        }
    } else {
        let left = pop_number(evaluator, &mut formula_error, &mut generated_error)?;
        let right = pop_number(evaluator, &mut formula_error, &mut generated_error)?;
        let mut sum = WideSum::default();
        if formula_error.is_none()
            && generated_error.is_none()
            && let (Some(left), Some(right)) = (left, right)
        {
            let result = match function {
                Function::SumX2MY2 => sum.push_difference_of_squares(left, right),
                Function::SumX2PY2 => sum.push_sum_of_squares(left, right),
                Function::SumXMY2 => sum.push_squared_difference(left, right),
                _ => {
                    return Err(EvaluationFailure::InvalidExpression(
                        "non-pair aggregate reached scalar pair dispatch",
                    ));
                },
            };
            if let Err(error) = result {
                remember(&mut generated_error, error);
            }
        }
        return finish_wide_scalar(evaluator, sum, formula_error, generated_error);
    }

    if let Some(error) = formula_error {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(failure) = span_failure {
        return Err(failure);
    }
    evaluator.push_value(WorkingValue::Error(
        generated_error.unwrap_or(ScalarError::Number),
    ))
}

fn pop_number<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    formula_error: &mut Option<ScalarError>,
    generated_error: &mut Option<ScalarError>,
) -> EvaluationResult<Option<f64>> {
    evaluator.charge_work(1)?;
    match evaluator.pop_value()? {
        WorkingValue::Error(error) => {
            remember(formula_error, error);
            Ok(None)
        },
        value => match super::to_number(value, evaluator)? {
            Ok(value) => Ok(Some(value)),
            Err(error) => {
                remember(generated_error, error);
                Ok(None)
            },
        },
    }
}

fn finish_numeric_scalar<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    aggregate: NumericAggregate,
    formula_error: Option<ScalarError>,
    generated_error: Option<ScalarError>,
) -> EvaluationResult<()> {
    if let Some(error) = formula_error.or(generated_error) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    evaluator.charge_work(1)?;
    evaluator.push_value(match aggregate.result() {
        Ok(value) => WorkingValue::Number(value.unwrap_or(0.0)),
        Err(error) => WorkingValue::Error(error),
    })
}

fn finish_wide_scalar<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    sum: WideSum,
    formula_error: Option<ScalarError>,
    generated_error: Option<ScalarError>,
) -> EvaluationResult<()> {
    if let Some(error) = formula_error.or(generated_error) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    evaluator.charge_work(1)?;
    evaluator.push_value(match sum.result() {
        Ok(value) => WorkingValue::Number(value.unwrap_or(0.0)),
        Err(error) => WorkingValue::Error(error),
    })
}

fn note_product_result(
    result: Result<(), ProductSumError>,
    generated_error: &mut Option<ScalarError>,
    span_failure: &mut Option<EvaluationFailure>,
) {
    match result {
        Ok(()) => {},
        Err(ProductSumError::Number) => remember(generated_error, ScalarError::Number),
        Err(ProductSumError::ExponentSpan { observed, limit }) => {
            if span_failure.is_none() {
                *span_failure = Some(super::local_limit(Resource::Memory, observed, limit));
            }
        },
    }
}

fn remember(slot: &mut Option<ScalarError>, error: ScalarError) {
    if slot.is_none() {
        *slot = Some(error);
    }
}

fn reverse_value_tail<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    count: usize,
) -> EvaluationResult<()> {
    evaluator.charge_work(u64::try_from(count).unwrap_or(u64::MAX))?;
    let start =
        evaluator
            .values
            .len()
            .checked_sub(count)
            .ok_or(EvaluationFailure::InvalidExpression(
                "aggregate value stack underflow",
            ))?;
    evaluator.values[start..].reverse();
    Ok(())
}

#[cfg(test)]
mod tests {
    use crate::codec::formula::evaluation::{
        EvaluationLimits, ScalarValue, evaluate_scalar_with_context,
    };
    use crate::codec::formula::expression::Expression;
    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Profile,
    };
    use std::{
        num::{NonZeroU64, NonZeroUsize},
        sync::Arc,
    };

    fn execution() -> ExecutionContext {
        let (_source, token) = CancellationSource::pair();
        ExecutionContext::new(
            Budget::root(
                Arc::from("ods-aggregate-scalar-test"),
                Limits::for_profile(Profile::Desktop),
            ),
            token,
            ExecutionLimits::new(
                NonZeroUsize::MIN,
                NonZeroUsize::MIN,
                NonZeroU64::new(1 << 30).unwrap(),
                1 << 20,
            )
            .unwrap(),
        )
    }

    fn evaluate(source: &str) -> Result<f64, super::super::EvaluationFailure> {
        let execution = execution();
        let expression = Expression::parse(source).unwrap();
        evaluate_scalar_with_context(&expression, &execution, &EvaluationLimits::default()).map(
            |value| match value.value() {
                ScalarValue::Number(value) => *value,
                other => panic!("unexpected aggregate scalar result: {other:?}"),
            },
        )
    }

    #[test]
    fn scalar_sequence_functions_convert_and_fold() {
        assert_eq!(evaluate("=SUM(1;2;3)").unwrap(), 6.0);
        assert_eq!(evaluate("=PRODUCT(2;3;4)").unwrap(), 24.0);
        assert_eq!(evaluate("=SUMSQ(2;3)").unwrap(), 13.0);
        assert_eq!(evaluate("=SUM(\"2\";TRUE())").unwrap(), 3.0);
    }

    #[test]
    fn scalar_forced_array_functions_use_the_1x1_shape() {
        assert_eq!(evaluate("=SUMPRODUCT(2;3)").unwrap(), 6.0);
        assert_eq!(evaluate("=SUMX2MY2(3;2)").unwrap(), 5.0);
        assert_eq!(evaluate("=SUMX2PY2(3;2)").unwrap(), 13.0);
        assert_eq!(evaluate("=SUMXMY2(2;3)").unwrap(), 1.0);
    }
}
