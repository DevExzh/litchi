//! Private OpenFormula Criterion compilation and matching.
//!
//! Database functions and the conditional aggregate family use the same
//! Criterion rules from OpenFormula §4.11.8.  Keeping the matcher here avoids
//! making either family approximate the other while leaving their distinct
//! range geometry and aggregation policies in their owning modules.

use super::{
    CellRead, EvaluationFailure, EvaluationResult, Resolver, RuntimeElement, RuntimeValue,
    ScalarError, ValueEvaluator, WorkingValue,
};

/// The scalar values that can participate in a Criterion comparison.
///
/// The value is intentionally borrowed.  Compiled text criteria therefore
/// retain the source text for the duration of one formula evaluation instead
/// of allocating a copy for each row or cell comparison.
#[derive(Clone, Copy, Debug)]
pub(super) enum CriterionValue<'a> {
    Empty,
    Number(f64),
    Logical(bool),
    Text(&'a str),
    Error(ScalarError),
    Complex,
}

impl<'a> CriterionValue<'a> {
    pub(super) fn from_element(element: &'a RuntimeElement<'_>) -> CriterionValue<'a> {
        match element {
            RuntimeElement::Empty => CriterionValue::Empty,
            RuntimeElement::Missing => CriterionValue::Error(ScalarError::NotAvailable),
            RuntimeElement::Present(WorkingValue::Number(value)) => CriterionValue::Number(*value),
            RuntimeElement::Present(WorkingValue::Logical(value)) => {
                CriterionValue::Logical(*value)
            },
            RuntimeElement::Present(WorkingValue::Text(value)) => {
                CriterionValue::Text(value.text.as_ref())
            },
            RuntimeElement::Present(WorkingValue::Error(error)) => CriterionValue::Error(*error),
            RuntimeElement::Present(WorkingValue::Complex(_)) => CriterionValue::Complex,
        }
    }

    pub(super) fn from_runtime_value(value: &'a RuntimeValue<'_>) -> CriterionValue<'a> {
        match value {
            RuntimeValue::Empty | RuntimeValue::Missing => CriterionValue::Empty,
            RuntimeValue::Scalar(WorkingValue::Number(value)) => CriterionValue::Number(*value),
            RuntimeValue::Scalar(WorkingValue::Logical(value)) => CriterionValue::Logical(*value),
            RuntimeValue::Scalar(WorkingValue::Text(value)) => {
                CriterionValue::Text(value.text.as_ref())
            },
            RuntimeValue::Scalar(WorkingValue::Error(error)) => CriterionValue::Error(*error),
            RuntimeValue::Scalar(WorkingValue::Complex(_)) => CriterionValue::Complex,
            RuntimeValue::Array(_) | RuntimeValue::Areas(_) | RuntimeValue::ScalarCell(_) => {
                CriterionValue::Error(ScalarError::Value)
            },
        }
    }

    pub(super) fn from_cell_read(read: CellRead<'a>) -> EvaluationResult<Self> {
        Ok(match read {
            CellRead::Empty => Self::Empty,
            CellRead::Number(value) if value.is_finite() => Self::Number(value),
            CellRead::Number(_) => Self::Error(ScalarError::Number),
            CellRead::Logical(value) => Self::Logical(value),
            CellRead::Text(value) => Self::Text(value),
            CellRead::Error(error) => Self::Error(error),
            CellRead::Unsupported => {
                return Err(EvaluationFailure::Unsupported(
                    super::super::UnsupportedKind::CellValue,
                ));
            },
        })
    }

    pub(super) fn is_empty(self) -> bool {
        matches!(self, Self::Empty)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum CriterionOperator {
    Equal,
    NotEqual,
    Less,
    LessEqual,
    Greater,
    GreaterEqual,
}

#[derive(Clone, Copy, Debug)]
pub(super) enum CriterionMatcher<'a> {
    Empty(CriterionOperator),
    Number(CriterionOperator, f64),
    Logical(CriterionOperator, bool),
    Text(CriterionOperator, &'a str),
}

/// Compile one Criterion value according to OpenFormula §4.11.8.
pub(super) fn compile_criterion<'source, 'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: CriterionValue<'source>,
) -> EvaluationResult<Result<CriterionMatcher<'source>, ScalarError>>
where
    R: Resolver + ?Sized,
{
    match value {
        CriterionValue::Empty => Ok(Ok(CriterionMatcher::Number(CriterionOperator::Equal, 0.0))),
        CriterionValue::Number(value) => Ok(Ok(CriterionMatcher::Number(
            CriterionOperator::Equal,
            value,
        ))),
        CriterionValue::Logical(value) => Ok(Ok(CriterionMatcher::Logical(
            CriterionOperator::Equal,
            value,
        ))),
        CriterionValue::Error(error) => Ok(Err(error)),
        CriterionValue::Complex => Ok(Err(ScalarError::Value)),
        CriterionValue::Text(text) => {
            evaluator.scalar.charge_bytes(text.len())?;
            let (operator, rhs, has_operator) = criterion_parts(text);
            if rhs.is_empty() {
                return Ok(Ok(CriterionMatcher::Empty(operator)));
            }
            if has_operator {
                if let Ok(number) = fast_float2::parse::<f64, _>(rhs) {
                    if number.is_finite() {
                        return Ok(Ok(CriterionMatcher::Number(operator, number)));
                    }
                }
            }
            Ok(Ok(CriterionMatcher::Text(operator, rhs)))
        },
    }
}

/// Match one candidate cell against one compiled Criterion.
pub(super) fn criterion_matches<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    matcher: CriterionMatcher<'_>,
    candidate: CriterionValue<'_>,
) -> EvaluationResult<Result<bool, ScalarError>>
where
    R: Resolver + ?Sized,
{
    if let CriterionValue::Error(error) = candidate {
        return Ok(Err(error));
    }
    let result = match matcher {
        CriterionMatcher::Empty(operator) => match operator {
            CriterionOperator::Equal => candidate.is_empty(),
            CriterionOperator::NotEqual => !candidate.is_empty(),
            CriterionOperator::Less
            | CriterionOperator::LessEqual
            | CriterionOperator::Greater
            | CriterionOperator::GreaterEqual => false,
        },
        CriterionMatcher::Number(operator, expected) => match candidate {
            CriterionValue::Number(actual) => compare_numbers(operator, actual, expected),
            CriterionValue::Empty
            | CriterionValue::Logical(_)
            | CriterionValue::Text(_)
            | CriterionValue::Complex => matches!(operator, CriterionOperator::NotEqual),
            CriterionValue::Error(_) => unreachable!("formula errors are handled above"),
        },
        CriterionMatcher::Logical(operator, expected) => match candidate {
            CriterionValue::Logical(actual) => compare_booleans(operator, actual, expected),
            CriterionValue::Empty
            | CriterionValue::Number(_)
            | CriterionValue::Text(_)
            | CriterionValue::Complex => matches!(operator, CriterionOperator::NotEqual),
            CriterionValue::Error(_) => unreachable!("formula errors are handled above"),
        },
        CriterionMatcher::Text(operator, expected) => match candidate {
            CriterionValue::Text(actual) => {
                evaluator.scalar.charge_bytes(
                    actual
                        .len()
                        .checked_add(expected.len())
                        .unwrap_or(usize::MAX),
                )?;
                compare_text(operator, actual, expected)
            },
            CriterionValue::Empty
            | CriterionValue::Number(_)
            | CriterionValue::Logical(_)
            | CriterionValue::Complex => matches!(operator, CriterionOperator::NotEqual),
            CriterionValue::Error(_) => unreachable!("formula errors are handled above"),
        },
    };
    Ok(Ok(result))
}

fn criterion_parts(value: &str) -> (CriterionOperator, &str, bool) {
    if let Some(rest) = value.strip_prefix(">=") {
        (CriterionOperator::GreaterEqual, rest, true)
    } else if let Some(rest) = value.strip_prefix("<=") {
        (CriterionOperator::LessEqual, rest, true)
    } else if let Some(rest) = value.strip_prefix("<>") {
        (CriterionOperator::NotEqual, rest, true)
    } else if let Some(rest) = value.strip_prefix('>') {
        (CriterionOperator::Greater, rest, true)
    } else if let Some(rest) = value.strip_prefix('<') {
        (CriterionOperator::Less, rest, true)
    } else if let Some(rest) = value.strip_prefix('=') {
        (CriterionOperator::Equal, rest, true)
    } else {
        (CriterionOperator::Equal, value, false)
    }
}

fn compare_numbers(operator: CriterionOperator, left: f64, right: f64) -> bool {
    match operator {
        CriterionOperator::Equal => left == right,
        CriterionOperator::NotEqual => left != right,
        CriterionOperator::Less => left < right,
        CriterionOperator::LessEqual => left <= right,
        CriterionOperator::Greater => left > right,
        CriterionOperator::GreaterEqual => left >= right,
    }
}

fn compare_booleans(operator: CriterionOperator, left: bool, right: bool) -> bool {
    compare_numbers(operator, u8::from(left) as f64, u8::from(right) as f64)
}

fn compare_text(operator: CriterionOperator, left: &str, right: &str) -> bool {
    let ordering = left.cmp(right);
    match operator {
        CriterionOperator::Equal => ordering == std::cmp::Ordering::Equal,
        CriterionOperator::NotEqual => ordering != std::cmp::Ordering::Equal,
        CriterionOperator::Less => ordering == std::cmp::Ordering::Less,
        CriterionOperator::LessEqual => ordering != std::cmp::Ordering::Greater,
        CriterionOperator::Greater => ordering == std::cmp::Ordering::Greater,
        CriterionOperator::GreaterEqual => ordering != std::cmp::Ordering::Less,
    }
}
