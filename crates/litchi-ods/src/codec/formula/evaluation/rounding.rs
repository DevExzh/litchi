//! OpenFormula 1.4 section 6.17 rounding functions.
//!
//! The kernels in this module operate on finite scalar `f64` values and do
//! not allocate.  Multiple rounding uses a remainder rather than a rounded
//! quotient: a quotient can lose the low bits which decide a tie even when
//! the final multiple is representable. Integer decimal digit counts quantize
//! the input's shortest round-trip decimal representation with checked integer
//! coefficients; fractional digits use checked binary scales. Non-finite
//! results become the evaluator's typed Number error.

use std::fmt;
use std::fmt::Write as _;

use super::{EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, WorkingValue};

#[derive(Clone, Copy)]
enum Function {
    Ceiling,
    Int,
    Floor,
    Mround,
    Round,
    RoundDown,
    RoundUp,
    Trunc,
}

fn function(name: &str) -> Option<Function> {
    if name.eq_ignore_ascii_case("CEILING") {
        Some(Function::Ceiling)
    } else if name.eq_ignore_ascii_case("INT") {
        Some(Function::Int)
    } else if name.eq_ignore_ascii_case("FLOOR") {
        Some(Function::Floor)
    } else if name.eq_ignore_ascii_case("MROUND") {
        Some(Function::Mround)
    } else if name.eq_ignore_ascii_case("ROUND") {
        Some(Function::Round)
    } else if name.eq_ignore_ascii_case("ROUNDDOWN") {
        Some(Function::RoundDown)
    } else if name.eq_ignore_ascii_case("ROUNDUP") {
        Some(Function::RoundUp)
    } else if name.eq_ignore_ascii_case("TRUNC") {
        Some(Function::Trunc)
    } else {
        None
    }
}

pub(super) fn is_rounding_function(name: &str) -> bool {
    function(name).is_some()
}

/// Apply a section 6.17 function in the scalar evaluator.
///
/// The direct evaluator represents an omitted optional argument as a Missing
/// node.  The value evaluator converts a referenced Empty cell to Number zero
/// before entering this scalar kernel; only a syntactic Missing node receives
/// the function-defined default.  A mandatory missing number remains
/// `#VALUE!`.
#[inline(never)]
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    match function(name) {
        Some(Function::Ceiling) => apply_ceiling_floor(evaluator, node, true),
        Some(Function::Int) => apply_int(evaluator, node),
        Some(Function::Floor) => apply_ceiling_floor(evaluator, node, false),
        Some(Function::Mround) => apply_mround(evaluator, node),
        Some(Function::Round) => apply_decimal(evaluator, node, DecimalMode::Nearest),
        Some(Function::RoundDown) => apply_decimal(evaluator, node, DecimalMode::TowardZero),
        Some(Function::RoundUp) => apply_decimal(evaluator, node, DecimalMode::AwayFromZero),
        Some(Function::Trunc) => apply_decimal(evaluator, node, DecimalMode::TowardZero),
        None => Err(EvaluationFailure::InvalidExpression(
            "unknown rounding function reached evaluator",
        )),
    }
}

fn formula_error(value: &WorkingValue<'_>) -> Option<ScalarError> {
    match value {
        WorkingValue::Error(error) => Some(*error),
        _ => None,
    }
}

fn push_result(
    evaluator: &mut Evaluator<'_, '_, '_>,
    result: Result<f64, ScalarError>,
) -> EvaluationResult<()> {
    let value = match result {
        Ok(value) if value.is_finite() => {
            // The profile canonicalizes both signs of zero at the formula
            // boundary.  This also keeps directed negative rounding from
            // leaking a `-0.0` through a scalar or matrix result.
            WorkingValue::Number(if value == 0.0 { 0.0 } else { value })
        },
        Ok(_) => WorkingValue::Error(ScalarError::Number),
        Err(error) => WorkingValue::Error(error),
    };
    evaluator.push_value(value)
}

fn optional_empty(node: Node<'_>, index: usize) -> bool {
    index >= node.child_count() || node.child(index).is_some_and(|child| child.is_missing())
}

fn optional_number<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    value: WorkingValue<'a>,
    is_empty: bool,
    default: f64,
) -> EvaluationResult<Result<f64, ScalarError>> {
    if is_empty {
        Ok(Ok(default))
    } else {
        super::to_number(value, evaluator)
    }
}

fn optional_integer<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    value: WorkingValue<'a>,
    is_empty: bool,
) -> EvaluationResult<Result<f64, ScalarError>> {
    if is_empty {
        Ok(Ok(0.0))
    } else {
        super::to_integer(value, evaluator)
    }
}

fn apply_int<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    if node.child_count() != 1 {
        return evaluator.finish_invalid_arity(node);
    }
    let value = evaluator.pop_value()?;
    if let Some(error) = formula_error(&value) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let value = match super::to_number(value, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    push_result(evaluator, Ok(value.floor()))
}

fn apply_ceiling_floor<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    ceiling: bool,
) -> EvaluationResult<()> {
    let count = node.child_count();
    if !(1..=3).contains(&count) {
        return evaluator.finish_invalid_arity(node);
    }

    let mode = if count == 3 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let significance = if count >= 2 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let number = evaluator.pop_value()?;
    let significance_empty = optional_empty(node, 1);
    let mode_empty = optional_empty(node, 2);

    // Check every supplied formula error before attempting a conversion.  A
    // later formula error must not be hidden by an earlier bad text number.
    if let Some(error) = formula_error(&number) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(value) = &significance {
        if !significance_empty {
            if let Some(error) = formula_error(value) {
                return evaluator.push_value(WorkingValue::Error(error));
            }
        }
    }
    if let Some(value) = &mode {
        if !mode_empty {
            if let Some(error) = formula_error(value) {
                return evaluator.push_value(WorkingValue::Error(error));
            }
        }
    }

    let number = match super::to_number(number, evaluator)? {
        Ok(value) if value.is_finite() => value,
        Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let significance = match significance {
        Some(value) => match optional_number(
            evaluator,
            value,
            significance_empty,
            if number < 0.0 { -1.0 } else { 1.0 },
        )? {
            Ok(value) if value.is_finite() => value,
            Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
            Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
        },
        None => {
            if number < 0.0 {
                -1.0
            } else {
                1.0
            }
        },
    };
    let mode = match mode {
        Some(value) => match optional_number(evaluator, value, mode_empty, 0.0)? {
            Ok(value) if value.is_finite() => value,
            Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
            Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
        },
        None => 0.0,
    };

    push_result(
        evaluator,
        ceiling_floor(number, significance, mode, ceiling),
    )
}

#[derive(Clone, Copy)]
enum MultipleDirection {
    Floor,
    Ceiling,
}

fn finite_add(left: f64, right: f64) -> Result<f64, ScalarError> {
    let result = left + right;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ScalarError::Number)
    }
}

/// Round a non-negative finite value to a positive finite multiple.  The
/// remainder path works even when `value / step` is too large to be finite.
fn absolute_multiple(
    value: f64,
    step: f64,
    direction: MultipleDirection,
) -> Result<f64, ScalarError> {
    if value == 0.0 {
        return Ok(0.0);
    }
    if !step.is_finite() || step <= 0.0 {
        return Err(ScalarError::Number);
    }
    let remainder = value % step;
    if remainder == 0.0 {
        return Ok(value);
    }
    match direction {
        MultipleDirection::Floor => Ok(value - remainder),
        MultipleDirection::Ceiling => finite_add(value, step - remainder),
    }
}

fn apply_sign(magnitude: f64, negative: bool) -> Result<f64, ScalarError> {
    let value = if negative { -magnitude } else { magnitude };
    if value.is_finite() {
        Ok(value)
    } else {
        Err(ScalarError::Number)
    }
}

fn ceiling_floor(
    number: f64,
    significance: f64,
    mode: f64,
    ceiling: bool,
) -> Result<f64, ScalarError> {
    if number == 0.0 || significance == 0.0 {
        return Ok(0.0);
    }
    if !number.is_finite() || !significance.is_finite() || !mode.is_finite() {
        return Err(ScalarError::Number);
    }

    let number_negative = number < 0.0;
    let significance_negative = significance < 0.0;
    if number_negative != significance_negative {
        return Err(ScalarError::Value);
    }

    let direction = if mode != 0.0 {
        if ceiling {
            MultipleDirection::Ceiling
        } else {
            MultipleDirection::Floor
        }
    } else if ceiling == (significance > 0.0) {
        MultipleDirection::Ceiling
    } else {
        MultipleDirection::Floor
    };
    let magnitude = absolute_multiple(number.abs(), significance.abs(), direction)?;
    apply_sign(magnitude, number_negative)
}

fn apply_mround<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    if node.child_count() != 2 {
        return evaluator.finish_invalid_arity(node);
    }
    let multiple = evaluator.pop_value()?;
    let number = evaluator.pop_value()?;
    if let Some(error) = formula_error(&number) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(error) = formula_error(&multiple) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let number = match super::to_number(number, evaluator)? {
        Ok(value) if value.is_finite() => value,
        Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let multiple = match super::to_number(multiple, evaluator)? {
        Ok(value) if value.is_finite() => value,
        Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    push_result(evaluator, mround(number, multiple))
}

fn mround(number: f64, multiple: f64) -> Result<f64, ScalarError> {
    if multiple == 0.0 {
        // Section 6.17.4 supplies no zero-multiple special case.  Its
        // defining quotient X/B is undefined, so this profile reports the
        // existing division-by-zero formula error.
        return Err(ScalarError::DivisionByZero);
    }
    if number == 0.0 {
        return Ok(0.0);
    }
    let step = multiple.abs();
    if !step.is_finite() {
        return Err(ScalarError::Number);
    }
    let magnitude = number.abs();
    let remainder = magnitude % step;
    if remainder == 0.0 {
        return Ok(number);
    }
    let lower = magnitude - remainder;
    // Choose the nearest multiple before constructing the upper candidate.
    // On a tie OpenFormula selects the greater numerical value, which is the
    // upper magnitude for positive inputs and the lower magnitude for
    // negative inputs.  If upper is selected but is outside the finite
    // Number domain, propagate #NUM!; lower is not a substitute for a nearer
    // mathematical result.
    let upper_is_closer =
        remainder > step - remainder || (remainder == step - remainder && number > 0.0);
    let magnitude = if upper_is_closer {
        finite_add(magnitude, step - remainder)?
    } else {
        lower
    };
    apply_sign(magnitude, number < 0.0)
}

#[derive(Clone, Copy)]
enum DecimalMode {
    Nearest,
    TowardZero,
    AwayFromZero,
}

fn apply_decimal<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    mode: DecimalMode,
) -> EvaluationResult<()> {
    let count = node.child_count();
    if !(1..=2).contains(&count) {
        return evaluator.finish_invalid_arity(node);
    }
    let digits = if count == 2 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let number = evaluator.pop_value()?;
    let digits_empty = optional_empty(node, 1);

    if let Some(error) = formula_error(&number) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(value) = &digits {
        if !digits_empty {
            if let Some(error) = formula_error(value) {
                return evaluator.push_value(WorkingValue::Error(error));
            }
        }
    }

    let number = match super::to_number(number, evaluator)? {
        Ok(value) if value.is_finite() => value,
        Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let digits = match digits {
        Some(value) if matches!(mode, DecimalMode::Nearest) => {
            match optional_number(evaluator, value, digits_empty, 0.0)? {
                Ok(value) if value.is_finite() => value,
                Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
                Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
            }
        },
        Some(value) => match optional_integer(evaluator, value, digits_empty)? {
            Ok(value) if value.is_finite() => value,
            Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
            Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
        },
        None => 0.0,
    };
    push_result(evaluator, decimal_round(number, digits, mode))
}

// A shortest binary64 significand has at most 17 digits. Scientific notation
// adds a point, sign, exponent marker and at most three exponent digits.
const DECIMAL_TEXT_CAPACITY: usize = 32;

struct FixedText<const N: usize> {
    bytes: [u8; N],
    length: usize,
}

impl<const N: usize> FixedText<N> {
    fn new() -> Self {
        Self {
            bytes: [0; N],
            length: 0,
        }
    }

    fn as_str(&self) -> Option<&str> {
        std::str::from_utf8(&self.bytes[..self.length]).ok()
    }

    fn push_byte(&mut self, byte: u8) -> Option<()> {
        if self.length == N {
            return None;
        }
        self.bytes[self.length] = byte;
        self.length += 1;
        Some(())
    }
}

impl<const N: usize> fmt::Write for FixedText<N> {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        if value.len() > N - self.length {
            return Err(fmt::Error);
        }
        let end = self.length + value.len();
        self.bytes[self.length..end].copy_from_slice(value.as_bytes());
        self.length = end;
        Ok(())
    }
}

#[derive(Clone, Copy)]
struct DecimalParts {
    coefficient: u64,
    exponent: i32,
}

/// Parse Rust's allocation-free shortest scientific display into `coefficient ×
/// 10^exponent`. The parser accepts fixed or exponent notation; the
/// significant coefficient of a finite binary64 value has at most 17 digits
/// after trailing zeroes are removed.
fn decimal_parts(number: f64) -> Option<DecimalParts> {
    let mut text = FixedText::<DECIMAL_TEXT_CAPACITY>::new();
    write!(&mut text, "{:e}", number.abs()).ok()?;
    let bytes = &text.bytes[..text.length];
    let mut digits = [0_u8; DECIMAL_TEXT_CAPACITY];
    let mut digit_count = 0_usize;
    let mut decimal_position = 0_i32;
    let mut after_decimal = false;
    let mut index = 0_usize;

    while index < bytes.len() {
        match bytes[index] {
            b'0'..=b'9' => {
                if digit_count == digits.len() {
                    return None;
                }
                digits[digit_count] = bytes[index] - b'0';
                digit_count += 1;
                if !after_decimal {
                    decimal_position = decimal_position.checked_add(1)?;
                }
                index += 1;
            },
            b'.' => {
                if after_decimal {
                    return None;
                }
                after_decimal = true;
                index += 1;
            },
            b'e' | b'E' => {
                index += 1;
                let negative = bytes.get(index) == Some(&b'-');
                if negative || bytes.get(index) == Some(&b'+') {
                    index += 1;
                }
                if index == bytes.len() {
                    return None;
                }
                let mut exponent = 0_i32;
                while index < bytes.len() {
                    let digit = bytes[index].checked_sub(b'0')?;
                    if digit > 9 {
                        return None;
                    }
                    exponent = exponent.checked_mul(10)?.checked_add(i32::from(digit))?;
                    index += 1;
                }
                decimal_position = if negative {
                    decimal_position.checked_sub(exponent)?
                } else {
                    decimal_position.checked_add(exponent)?
                };
            },
            _ => return None,
        }
    }
    if digit_count == 0 {
        return None;
    }

    let mut first = 0_usize;
    while first < digit_count && digits[first] == 0 {
        first += 1;
    }
    if first == digit_count {
        return Some(DecimalParts {
            coefficient: 0,
            exponent: 0,
        });
    }

    let mut last = digit_count;
    let mut exponent = decimal_position.checked_sub(digit_count as i32)?;
    while last > first + 1 && digits[last - 1] == 0 {
        last -= 1;
        exponent = exponent.checked_add(1)?;
    }
    let mut coefficient = 0_u64;
    for &digit in &digits[first..last] {
        coefficient = coefficient.checked_mul(10)?.checked_add(u64::from(digit))?;
    }
    Some(DecimalParts {
        coefficient,
        exponent,
    })
}

fn encode_decimal(coefficient: u64, exponent: i32, negative: bool) -> Option<f64> {
    if coefficient == 0 {
        return Some(0.0);
    }
    let mut digits = [0_u8; 20];
    let mut length = 0_usize;
    let mut value = coefficient;
    while value != 0 {
        digits[length] = b'0' + u8::try_from(value % 10).ok()?;
        length += 1;
        value /= 10;
    }
    digits[..length].reverse();

    let mut text = FixedText::<DECIMAL_TEXT_CAPACITY>::new();
    if negative {
        text.push_byte(b'-')?;
    }
    text.push_byte(digits[0])?;
    if length > 1 {
        text.push_byte(b'.')?;
        for &digit in &digits[1..length] {
            text.push_byte(digit)?;
        }
    }
    let scientific_exponent = exponent.checked_add(length as i32 - 1)?;
    write!(&mut text, "e{scientific_exponent}").ok()?;
    let text = text.as_str()?;
    let value = fast_float2::parse::<f64, _>(text).ok()?;
    Some(value)
}

/// Quantize an integer `Digits` argument using the shortest decimal value of
/// the finite binary64 input.  This keeps distinct adjacent values distinct
/// while treating ordinary literals such as `0.3`, `0.07`, and `1.15` as their
/// displayed decimal values. A failed checked conversion is reported rather
/// than silently substituting a different numerical policy.
fn decimal_integer_round(
    number: f64,
    digits: i32,
    mode: DecimalMode,
) -> Option<Result<f64, ScalarError>> {
    let parts = decimal_parts(number)?;
    let shift = parts.exponent.checked_add(digits)?;
    if shift >= 0 {
        return Some(Ok(number));
    }

    let cut = shift.checked_neg()? as u32;
    let (mut quotient, remainder_nonzero, rounds_up) = if cut > 19 {
        (0_u64, parts.coefficient != 0, false)
    } else {
        let divisor = 10_u64.pow(cut);
        let quotient = parts.coefficient / divisor;
        let remainder = parts.coefficient % divisor;
        let rounds_up = matches!(mode, DecimalMode::Nearest) && remainder >= divisor.div_ceil(2);
        (quotient, remainder != 0, rounds_up)
    };
    if matches!(mode, DecimalMode::AwayFromZero) && remainder_nonzero {
        quotient = quotient.checked_add(1)?;
    } else if rounds_up {
        quotient = quotient.checked_add(1)?;
    }

    let result = encode_decimal(quotient, -digits, number.is_sign_negative())?;
    Some(if result.is_finite() {
        Ok(result)
    } else {
        Err(ScalarError::Number)
    })
}

/// Return the profile's finite `f64` representation of a decimal power.  The
/// binary path is used only for fractional Digits.
fn decimal_power(exponent: f64) -> f64 {
    10.0_f64.powf(exponent)
}

/// Use checked decimal scales for fractional Digits. A positive scale is formed
/// by multiplication where possible and a reciprocal division fallback keeps
/// subnormal inputs observable.  Negative Digits divide by the scale and
/// multiply the selected result back; an overflowing scale is classified from
/// the mathematical nearest candidates instead of being treated as zero.
fn decimal_binary_round(number: f64, digits: f64, mode: DecimalMode) -> Result<f64, ScalarError> {
    let magnitude_digits = digits.abs();
    let factor = decimal_power(magnitude_digits);
    let inverse = decimal_power(-magnitude_digits);
    let (quotient, multiply_result) = if digits >= 0.0 {
        let multiplied = number * factor;
        if multiplied.is_finite() && (multiplied != 0.0 || number == 0.0) {
            (multiplied, true)
        } else if inverse != 0.0 {
            let divided = number / inverse;
            if divided.is_finite() {
                (divided, false)
            } else {
                return Ok(number);
            }
        } else {
            return Ok(number);
        }
    } else {
        if factor == 0.0 {
            return Ok(number);
        }
        if !factor.is_finite() {
            return match mode {
                DecimalMode::TowardZero => Ok(0.0),
                DecimalMode::AwayFromZero => Err(ScalarError::Number),
                DecimalMode::Nearest => {
                    let input_exponent = number.abs().log10();
                    if input_exponent >= magnitude_digits - std::f64::consts::LOG10_2 {
                        Err(ScalarError::Number)
                    } else {
                        Ok(0.0)
                    }
                },
            };
        }
        let quotient = number / factor;
        if !quotient.is_finite() {
            return Ok(number);
        }
        (quotient, false)
    };
    let rounded = match mode {
        DecimalMode::Nearest => quotient.round(),
        DecimalMode::TowardZero => quotient.trunc(),
        DecimalMode::AwayFromZero => {
            if quotient < 0.0 {
                quotient.floor()
            } else {
                quotient.ceil()
            }
        },
    };
    let result = if multiply_result {
        rounded / factor
    } else if digits >= 0.0 {
        rounded * inverse
    } else {
        rounded * factor
    };
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ScalarError::Number)
    }
}

fn decimal_round(number: f64, digits: f64, mode: DecimalMode) -> Result<f64, ScalarError> {
    if !number.is_finite() || !digits.is_finite() {
        return Err(ScalarError::Number);
    }
    if number == 0.0 {
        return Ok(0.0);
    }
    if digits.fract() == 0.0 {
        // Outside this bounded integer range the mathematical quantum is
        // either finer than the smallest binary64 input or larger than every
        // finite binary64 input.  Resolve those cases before any float-to-int
        // cast and avoid asking `powf` to classify an arbitrary exponent.
        if digits > 324.0 {
            return Ok(number);
        }
        if digits < -309.0 {
            return match mode {
                DecimalMode::AwayFromZero => Err(ScalarError::Number),
                DecimalMode::Nearest | DecimalMode::TowardZero => Ok(0.0),
            };
        }
        // The fixed decimal path is exact for integer Digits in the range
        // where a finite f64 can carry the requested decimal quantum.  A
        // formatting or checked-arithmetic failure is a typed numeric error,
        // rather than a silent switch to a different numerical policy.
        return decimal_integer_round(number, digits as i32, mode)
            .unwrap_or(Err(ScalarError::Number));
    }
    decimal_binary_round(number, digits, mode)
}
