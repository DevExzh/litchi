//! OpenFormula 1.4 section 6.16 elementary real mathematical functions.
//!
//! The kernels in this module operate on finite `f64` values and do not
//! allocate.  The scalar evaluator and the value evaluator both enter this
//! dispatch, so a scalar, broadcast array, and projected reference have the
//! same domain and non-finite-result behavior.

use std::f64::consts::PI;

use super::{EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, WorkingValue};

#[derive(Clone, Copy)]
enum Function {
    Abs,
    Exp,
    Ln,
    Log,
    Log10,
    Power,
    Sqrt,
    SqrtPi,
    Sign,
    Mod,
    Quotient,
}

fn function(name: &str) -> Option<Function> {
    if name.eq_ignore_ascii_case("ABS") {
        Some(Function::Abs)
    } else if name.eq_ignore_ascii_case("EXP") {
        Some(Function::Exp)
    } else if name.eq_ignore_ascii_case("LN") {
        Some(Function::Ln)
    } else if name.eq_ignore_ascii_case("LOG") {
        Some(Function::Log)
    } else if name.eq_ignore_ascii_case("LOG10") {
        Some(Function::Log10)
    } else if name.eq_ignore_ascii_case("POWER") {
        Some(Function::Power)
    } else if name.eq_ignore_ascii_case("SQRT") {
        Some(Function::Sqrt)
    } else if name.eq_ignore_ascii_case("SQRTPI") {
        Some(Function::SqrtPi)
    } else if name.eq_ignore_ascii_case("SIGN") {
        Some(Function::Sign)
    } else if name.eq_ignore_ascii_case("MOD") {
        Some(Function::Mod)
    } else if name.eq_ignore_ascii_case("QUOTIENT") {
        Some(Function::Quotient)
    } else {
        None
    }
}

pub(super) fn is_elementary_function(name: &str) -> bool {
    function(name).is_some()
}

/// Apply one section 6.16 function after eager argument evaluation.
#[inline(never)]
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    match function(name) {
        Some(Function::Log) => apply_log(evaluator, node),
        Some(function @ (Function::Power | Function::Mod | Function::Quotient)) => {
            apply_binary(evaluator, node, function)
        },
        Some(function) => apply_unary(evaluator, node, function),
        None => Err(EvaluationFailure::InvalidExpression(
            "unknown elementary function reached evaluator",
        )),
    }
}

fn apply_unary<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    function: Function,
) -> EvaluationResult<()> {
    if node.child_count() != 1 {
        return evaluator.finish_invalid_arity(node);
    }
    let value = evaluator.pop_value()?;
    if let WorkingValue::Error(error) = value {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let value = match super::to_number(value, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    push_number(evaluator, unary_value(function, value))
}

fn apply_log<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    if !(1..=2).contains(&node.child_count()) {
        return evaluator.finish_invalid_arity(node);
    }

    // The evaluator visits arguments from left to right but its value stack
    // is popped from the right.  Check and convert the number first so a
    // formula error is retained in source order when both arguments fail.
    let base = if node.child_count() == 2 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let number = evaluator.pop_value()?;
    // Formula errors take precedence over conversion errors in any eagerly
    // evaluated argument.  This matters for e.g. LOG("bad";#N/A): the
    // explicit #N/A is retained instead of being hidden by #VALUE! from the
    // text conversion.
    if let WorkingValue::Error(error) = &number {
        return evaluator.push_value(WorkingValue::Error(*error));
    }
    if let Some(value) = &base {
        if let WorkingValue::Error(error) = value {
            return evaluator.push_value(WorkingValue::Error(*error));
        }
    }
    let number = match super::to_number(number, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let base = match base {
        Some(value) => match super::to_number(value, evaluator)? {
            Ok(value) => value,
            Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
        },
        None => 10.0,
    };
    push_number(evaluator, logarithm(number, base))
}

fn apply_binary<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    function: Function,
) -> EvaluationResult<()> {
    if node.child_count() != 2 {
        return evaluator.finish_invalid_arity(node);
    }
    let right = evaluator.pop_value()?;
    let left = evaluator.pop_value()?;
    if let WorkingValue::Error(error) = left {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let WorkingValue::Error(error) = right {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let left = match super::to_number(left, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let right = match super::to_number(right, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let result = match function {
        Function::Power => power_result(left, right),
        Function::Mod => modulo(left, right),
        Function::Quotient => quotient(left, right),
        _ => Err(ScalarError::Value),
    };
    push_number(evaluator, result)
}

fn unary_value(function: Function, value: f64) -> Result<f64, ScalarError> {
    match function {
        Function::Abs => Ok(value.abs()),
        Function::Exp => Ok(value.exp()),
        Function::Ln => domain_result(value, value > 0.0, f64::ln),
        Function::Log10 => domain_result(value, value > 0.0, f64::log10),
        Function::Sqrt => domain_result(value, value >= 0.0, f64::sqrt),
        Function::SqrtPi => {
            if value < 0.0 {
                Err(ScalarError::Number)
            } else {
                // Multiplying first can overflow even though the requested
                // square root is representable.  Factoring the constant
                // keeps both the large and subnormal finite domains useful.
                Ok(value.sqrt() * PI.sqrt())
            }
        },
        Function::Sign => Ok(if value < 0.0 {
            -1.0
        } else if value > 0.0 {
            1.0
        } else {
            0.0
        }),
        Function::Power | Function::Log | Function::Mod | Function::Quotient => {
            Err(ScalarError::Value)
        },
    }
}

fn logarithm(number: f64, base: f64) -> Result<f64, ScalarError> {
    // The standard constrains the number to be positive.  A real logarithm
    // also needs a positive base other than one; invalid bases naturally
    // produce a typed Number error instead of exposing NaN or infinity.
    if number <= 0.0 || base <= 0.0 || base == 1.0 {
        return Err(ScalarError::Number);
    }
    Ok(number.ln() / base.ln())
}

/// Shared by the POWER function and the infix `^` operator.
pub(super) fn power_result(base: f64, exponent: f64) -> Result<f64, ScalarError> {
    if base == 0.0 {
        if exponent == 0.0 {
            return Ok(1.0);
        }
        if exponent < 0.0 {
            return Err(ScalarError::Number);
        }
    }
    let result = base.powf(exponent);
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ScalarError::Number)
    }
}

/// Compute a remainder without forming a quotient or subtracting a large
/// nearly-equal product.  The OpenFormula remainder has the divisor's sign.
fn modulo(dividend: f64, divisor: f64) -> Result<f64, ScalarError> {
    if divisor == 0.0 {
        return Err(ScalarError::DivisionByZero);
    }
    // `%` performs the bounded floating-point remainder directly.  Adjusting
    // a nonzero result by one divisor supplies the ODF sign-of-divisor rule
    // without ever materializing the potentially overflowing quotient.
    let remainder = dividend % divisor;
    if remainder == 0.0 {
        return Ok(0.0);
    }
    let remainder = if remainder.is_sign_negative() != divisor.is_sign_negative() {
        remainder + divisor
    } else {
        remainder
    };
    if remainder.is_finite() {
        Ok(remainder)
    } else {
        Err(ScalarError::Number)
    }
}

fn quotient(dividend: f64, divisor: f64) -> Result<f64, ScalarError> {
    if divisor == 0.0 {
        return Err(ScalarError::DivisionByZero);
    }
    let quotient = (dividend / divisor).trunc();
    if quotient.is_finite() {
        Ok(quotient)
    } else {
        Err(ScalarError::Number)
    }
}

fn domain_result(value: f64, valid: bool, function: fn(f64) -> f64) -> Result<f64, ScalarError> {
    if valid {
        Ok(function(value))
    } else {
        Err(ScalarError::Number)
    }
}

fn push_number(
    evaluator: &mut Evaluator<'_, '_, '_>,
    result: Result<f64, ScalarError>,
) -> EvaluationResult<()> {
    let value = match result {
        Ok(value) if value.is_finite() => WorkingValue::Number(value),
        Ok(_) => WorkingValue::Error(ScalarError::Number),
        Err(error) => WorkingValue::Error(error),
    };
    evaluator.push_value(value)
}
