//! OpenFormula 1.4 section 6.16 trigonometric and hyperbolic functions.
//!
//! The evaluator keeps these kernels scalar and allocation-free.  The value
//! evaluator invokes the same kernels once per broadcast cell, so domain and
//! non-finite handling remain identical for scalar, array, and reference
//! inputs.  Reciprocal hyperbolic functions use stable forms for large finite
//! arguments: evaluating `sinh` or `cosh` first can overflow even when the
//! reciprocal is still a representable (possibly subnormal) `f64`.

use std::f64::consts::PI;

use super::{EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, WorkingValue};

#[derive(Clone, Copy)]
enum Function {
    Acos,
    Acosh,
    Acot,
    Acoth,
    Asin,
    Asinh,
    Atan,
    Atan2,
    Atanh,
    Cos,
    Cosh,
    Cot,
    Coth,
    Csc,
    Csch,
    Degrees,
    Pi,
    Radians,
    Sec,
    Sech,
    Sin,
    Sinh,
    Tan,
    Tanh,
}

fn function(name: &str) -> Option<Function> {
    if name.eq_ignore_ascii_case("ACOS") {
        Some(Function::Acos)
    } else if name.eq_ignore_ascii_case("ACOSH") {
        Some(Function::Acosh)
    } else if name.eq_ignore_ascii_case("ACOT") {
        Some(Function::Acot)
    } else if name.eq_ignore_ascii_case("ACOTH") {
        Some(Function::Acoth)
    } else if name.eq_ignore_ascii_case("ASIN") {
        Some(Function::Asin)
    } else if name.eq_ignore_ascii_case("ASINH") {
        Some(Function::Asinh)
    } else if name.eq_ignore_ascii_case("ATAN") {
        Some(Function::Atan)
    } else if name.eq_ignore_ascii_case("ATAN2") {
        Some(Function::Atan2)
    } else if name.eq_ignore_ascii_case("ATANH") {
        Some(Function::Atanh)
    } else if name.eq_ignore_ascii_case("COS") {
        Some(Function::Cos)
    } else if name.eq_ignore_ascii_case("COSH") {
        Some(Function::Cosh)
    } else if name.eq_ignore_ascii_case("COT") {
        Some(Function::Cot)
    } else if name.eq_ignore_ascii_case("COTH") {
        Some(Function::Coth)
    } else if name.eq_ignore_ascii_case("CSC") {
        Some(Function::Csc)
    } else if name.eq_ignore_ascii_case("CSCH") {
        Some(Function::Csch)
    } else if name.eq_ignore_ascii_case("DEGREES") {
        Some(Function::Degrees)
    } else if name.eq_ignore_ascii_case("PI") {
        Some(Function::Pi)
    } else if name.eq_ignore_ascii_case("RADIANS") {
        Some(Function::Radians)
    } else if name.eq_ignore_ascii_case("SEC") {
        Some(Function::Sec)
    } else if name.eq_ignore_ascii_case("SECH") {
        Some(Function::Sech)
    } else if name.eq_ignore_ascii_case("SIN") {
        Some(Function::Sin)
    } else if name.eq_ignore_ascii_case("SINH") {
        Some(Function::Sinh)
    } else if name.eq_ignore_ascii_case("TAN") {
        Some(Function::Tan)
    } else if name.eq_ignore_ascii_case("TANH") {
        Some(Function::Tanh)
    } else {
        None
    }
}

pub(super) fn is_trigonometric_function(name: &str) -> bool {
    function(name).is_some()
}

/// Apply one section 6.16 function after eager argument evaluation.
///
/// Keeping the family in one outlined dispatch preserves the evaluator's
/// small scalar hot loop.  The numerical kernels below are plain functions:
/// they allocate no temporary storage and return a typed formula error when a
/// domain violation or non-finite result cannot be represented by the finite
/// Number profile.
#[inline(never)]
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    match function(name) {
        Some(Function::Atan2) => apply_atan2(evaluator, node),
        Some(Function::Pi) => apply_pi(evaluator, node),
        Some(function) => apply_unary(evaluator, node, function),
        None => Err(EvaluationFailure::InvalidExpression(
            "unknown trigonometric function reached evaluator",
        )),
    }
}

fn apply_pi<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    if node.child_count() != 0 {
        return evaluator.finish_invalid_arity(node);
    }
    evaluator.push_value(WorkingValue::Number(PI))
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
    let result = unary_value(function, value);
    push_number(evaluator, result)
}

fn apply_atan2<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    if node.child_count() != 2 {
        return evaluator.finish_invalid_arity(node);
    }
    // OpenFormula names the coordinates ATAN2(x; y).  Rust's atan2 method is
    // atan2(y, x), hence the deliberate reversal at the call site below.
    let y = evaluator.pop_value()?;
    let x = evaluator.pop_value()?;
    if let WorkingValue::Error(error) = x {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let WorkingValue::Error(error) = y {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let x = match super::to_number(x, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let y = match super::to_number(y, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    if x == 0.0 && y == 0.0 {
        // Part 4 leaves ATAN2(0;0) implementation-defined.  A formula error
        // keeps the finite profile explicit instead of inventing a direction.
        return evaluator.push_value(WorkingValue::Error(ScalarError::Number));
    }
    let result = y.atan2(x);
    // The specified principal interval includes +pi and excludes -pi on the
    // exact negative-x, zero-y branch cut.  Keep a nonzero lower-quadrant y
    // value intact even when finite libm rounding places it at -pi.
    let result = if y == 0.0 && x < 0.0 { PI } else { result };
    push_number(evaluator, Ok(result))
}

fn unary_value(function: Function, value: f64) -> Result<f64, ScalarError> {
    match function {
        Function::Acos => domain_result(value, (-1.0..=1.0).contains(&value), f64::acos),
        Function::Acosh => {
            if value < 1.0 {
                Err(ScalarError::Number)
            } else if value < 2.0 {
                // Near one, this form avoids the cancellation in x² - 1.
                Ok(((value - 1.0) + (value - 1.0).sqrt() * (value + 1.0).sqrt()).ln_1p())
            } else {
                // x² can overflow even when acosh(x) is finite.  Factoring x
                // out of x + sqrt(x² - 1) keeps the computation finite.
                let inverse = 1.0 / value;
                Ok(value.ln() + (1.0 - inverse * inverse).sqrt().ln_1p())
            }
        },
        // atan2(1, x) selects the (0, pi) branch and remains positive for
        // large positive finite x, where pi/2 - atan(x) can round to zero.
        Function::Acot => Ok(1.0_f64.atan2(value)),
        Function::Acoth => {
            if value.abs() <= 1.0 {
                Err(ScalarError::Number)
            } else {
                // 1/2 ln((x+1)/(x-1)), written in terms of the positive
                // magnitude so values just below -1 do not round their
                // numerator/denominator quotient to -1.  `ln_1p` also keeps
                // the small result for very large finite inputs.
                let magnitude = value.abs();
                let result = 0.5 * (2.0 / (magnitude - 1.0)).ln_1p();
                Ok(result.copysign(value))
            }
        },
        Function::Asin => domain_result(value, (-1.0..=1.0).contains(&value), f64::asin),
        Function::Asinh => Ok(stable_asinh(value)),
        Function::Atan => Ok(value.atan()),
        Function::Atanh => domain_result(value, value > -1.0 && value < 1.0, f64::atanh),
        Function::Cos => Ok(value.cos()),
        Function::Cosh => Ok(value.cosh()),
        Function::Cot => reciprocal(value.tan()),
        Function::Coth => reciprocal(value.tanh()),
        Function::Csc => reciprocal(value.sin()),
        Function::Csch => reciprocal_sinh(value),
        Function::Degrees => Ok(value.to_degrees()),
        Function::Radians => Ok(value.to_radians()),
        Function::Sec => reciprocal(value.cos()),
        Function::Sech => reciprocal_cosh(value),
        Function::Sin => Ok(value.sin()),
        Function::Sinh => Ok(value.sinh()),
        Function::Tan => Ok(value.tan()),
        Function::Tanh => Ok(value.tanh()),
        Function::Atan2 | Function::Pi => Err(ScalarError::Value),
    }
}

fn domain_result(value: f64, valid: bool, function: fn(f64) -> f64) -> Result<f64, ScalarError> {
    if valid {
        Ok(function(value))
    } else {
        Err(ScalarError::Number)
    }
}

fn reciprocal(value: f64) -> Result<f64, ScalarError> {
    if value == 0.0 {
        Err(ScalarError::DivisionByZero)
    } else {
        Ok(1.0 / value)
    }
}

/// Return `1/sinh(value)` without first overflowing `sinh` for large values.
fn reciprocal_sinh(value: f64) -> Result<f64, ScalarError> {
    if value == 0.0 {
        return Err(ScalarError::DivisionByZero);
    }
    let magnitude = value.abs();
    let result = if magnitude < 20.0 {
        1.0 / value.sinh()
    } else {
        // Evaluating exp(-magnitude) first can underflow around 745 even
        // though 2*exp(-magnitude) is still the smallest subnormal.  Splitting
        // the exponent keeps both intermediate values normal/subnormal enough
        // for the final product to round correctly.
        let half = (-0.5 * magnitude).exp();
        let result = 2.0 * half * half;
        if value.is_sign_negative() {
            -result
        } else {
            result
        }
    };
    Ok(result)
}

/// Evaluate asinh without overflowing the square in its logarithmic form.
fn stable_asinh(value: f64) -> f64 {
    let magnitude = value.abs();
    if magnitude <= 1.0 {
        value.asinh()
    } else {
        let inverse = 1.0 / magnitude;
        let result = magnitude.ln() + (1.0 + inverse * inverse).sqrt().ln_1p();
        result.copysign(value)
    }
}

/// Return `1/cosh(value)` without first overflowing `cosh` for large values.
fn reciprocal_cosh(value: f64) -> Result<f64, ScalarError> {
    let magnitude = value.abs();
    Ok(if magnitude < 20.0 {
        1.0 / value.cosh()
    } else {
        let half = (-0.5 * magnitude).exp();
        2.0 * half * half
    })
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
