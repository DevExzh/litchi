//! OpenFormula 1.4 section 6.8 complex-number functions.
//!
//! The evaluator keeps complex values as a finite pair of `f64` components
//! and a representation suffix.  The suffix is metadata for text round trips;
//! it does not participate in arithmetic.  This module owns conversion,
//! formatting, and the numerical kernels so the scalar evaluator only needs
//! to schedule the already-eager arguments.

use super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, NumberText, ScalarError, TextValue,
    WorkingValue, ensure_capacity, local_limit,
};
use litchi_core::{Reservation, Resource};
use std::fmt::{self, Display, Formatter, Write as FmtWrite};

const COMPLEX_FUNCTION_NAMES: [&str; 26] = [
    "COMPLEX",
    "IMABS",
    "IMAGINARY",
    "IMARGUMENT",
    "IMCONJUGATE",
    "IMCOS",
    "IMCOSH",
    "IMCOT",
    "IMCSC",
    "IMCSCH",
    "IMDIV",
    "IMEXP",
    "IMLN",
    "IMLOG10",
    "IMLOG2",
    "IMPOWER",
    "IMPRODUCT",
    "IMREAL",
    "IMSIN",
    "IMSINH",
    "IMSEC",
    "IMSECH",
    "IMSQRT",
    "IMSUB",
    "IMSUM",
    "IMTAN",
];

mod numerics;

/// Construction failures for a public complex value.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// Either component is NaN or infinite.
    NonFinite,
    /// The representation suffix is not lowercase `i` or `j`.
    InvalidSuffix,
}

impl Display for Error {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::NonFinite => "complex component is not finite",
            Self::InvalidSuffix => "complex suffix must be lowercase i or j",
        })
    }
}

impl std::error::Error for Error {}

/// A bounded aggregate argument buffer whose storage reservation follows the
/// values until both have been consumed or dropped.  Fixed-arity functions
/// pop directly from the evaluator stack and do not create this buffer.
struct ComplexArguments<'a> {
    values: Vec<WorkingValue<'a>>,
    _reservation: Option<Reservation>,
}

struct ComplexArgumentsIntoIter<'a> {
    values: std::vec::IntoIter<WorkingValue<'a>>,
    _reservation: Option<Reservation>,
}

impl<'a> ComplexArguments<'a> {
    fn as_slice(&self) -> &[WorkingValue<'a>] {
        &self.values
    }

    fn into_iter(self) -> ComplexArgumentsIntoIter<'a> {
        ComplexArgumentsIntoIter {
            values: self.values.into_iter(),
            _reservation: self._reservation,
        }
    }
}

impl<'a> IntoIterator for ComplexArguments<'a> {
    type Item = WorkingValue<'a>;
    type IntoIter = ComplexArgumentsIntoIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        ComplexArguments::into_iter(self)
    }
}

impl<'a> Iterator for ComplexArgumentsIntoIter<'a> {
    type Item = WorkingValue<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        self.values.next()
    }
}

/// A finite OpenFormula complex number.
///
/// The suffix is either lowercase `i` or `j` and is retained by the value.
/// Text formatting includes it when the imaginary component is nonzero; a
/// real-only value formats as a real number. Equality compares components
/// exactly and ignores the suffix, as does arithmetic.  Use [`Complex::new`] to construct a checked value.
#[derive(Clone, Copy, Debug)]
pub struct Complex {
    real: f64,
    imaginary: f64,
    suffix: char,
}

impl PartialEq for Complex {
    fn eq(&self, other: &Self) -> bool {
        self.real == other.real && self.imaginary == other.imaginary
    }
}

impl Complex {
    /// Construct a finite complex number with an explicit lowercase suffix.
    pub fn new(real: f64, imaginary: f64, suffix: char) -> Result<Self, Error> {
        if !real.is_finite() || !imaginary.is_finite() {
            return Err(Error::NonFinite);
        }
        if !matches!(suffix, 'i' | 'j') {
            return Err(Error::InvalidSuffix);
        }
        Ok(Self {
            real,
            imaginary,
            suffix,
        })
    }

    /// Return the real coefficient.
    #[must_use]
    pub const fn real(self) -> f64 {
        self.real
    }

    /// Return the imaginary coefficient.
    #[must_use]
    pub const fn imaginary(self) -> f64 {
        self.imaginary
    }

    /// Return the lowercase representation suffix (`i` or `j`).
    #[must_use]
    pub const fn suffix(self) -> char {
        self.suffix
    }

    fn checked(real: f64, imaginary: f64, suffix: char) -> Result<Self, ScalarError> {
        Self::new(real, imaginary, suffix).map_err(|error| match error {
            Error::NonFinite => ScalarError::Number,
            Error::InvalidSuffix => ScalarError::Value,
        })
    }

    const fn zero(suffix: char) -> Self {
        Self {
            real: 0.0,
            imaginary: 0.0,
            suffix,
        }
    }

    const fn one(suffix: char) -> Self {
        Self {
            real: 1.0,
            imaginary: 0.0,
            suffix,
        }
    }

    pub(super) fn negate(self) -> Self {
        Self {
            real: -self.real,
            imaginary: -self.imaginary,
            suffix: self.suffix,
        }
    }

    fn is_zero(self) -> bool {
        self.real == 0.0 && self.imaginary == 0.0
    }
}

impl Display for Complex {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> fmt::Result {
        if self.imaginary == 0.0 {
            return write!(formatter, "{}", self.real);
        }
        if self.real == 0.0 {
            if self.imaginary == 1.0 {
                return write!(formatter, "{}", self.suffix);
            }
            if self.imaginary == -1.0 {
                return write!(formatter, "-{}", self.suffix);
            }
            return write!(formatter, "{}{}", self.imaginary, self.suffix);
        }

        write!(formatter, "{}", self.real)?;
        if self.imaginary > 0.0 {
            write!(formatter, "+")?;
        }
        if self.imaginary == 1.0 {
            write!(formatter, "{}", self.suffix)
        } else if self.imaginary == -1.0 {
            write!(formatter, "-{}", self.suffix)
        } else {
            write!(formatter, "{}{}", self.imaginary, self.suffix)
        }
    }
}

/// Return whether `name` belongs to the complete section 6.8 family.
pub(super) fn is_complex_function(name: &str) -> bool {
    COMPLEX_FUNCTION_NAMES
        .iter()
        .any(|candidate| name.eq_ignore_ascii_case(candidate))
}

/// Apply one section 6.8 function after its eager arguments are on the scalar
/// value stack.  The function family is outlined to keep the ordinary scalar
/// evaluator's hot dispatch small.
#[inline(never)]
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    if name.eq_ignore_ascii_case("COMPLEX") {
        return apply_complex(evaluator, node);
    }
    if name.eq_ignore_ascii_case("IMABS") {
        return apply_unary_number(evaluator, node, unary_abs);
    }
    if name.eq_ignore_ascii_case("IMAGINARY") {
        return apply_unary_number(evaluator, node, |value| Ok(value.imaginary));
    }
    if name.eq_ignore_ascii_case("IMARGUMENT") {
        return apply_unary_number(evaluator, node, |value| Ok(complex_argument(value)));
    }
    if name.eq_ignore_ascii_case("IMCONJUGATE") {
        return apply_unary_complex(evaluator, node, |value| {
            Ok(Complex {
                real: value.real,
                imaginary: -value.imaginary,
                suffix: value.suffix,
            })
        });
    }
    if name.eq_ignore_ascii_case("IMCOS") {
        return apply_unary_complex(evaluator, node, numerics::cos);
    }
    if name.eq_ignore_ascii_case("IMCOSH") {
        return apply_unary_complex(evaluator, node, numerics::cosh);
    }
    if name.eq_ignore_ascii_case("IMCOT") {
        return apply_unary_complex(evaluator, node, numerics::cot);
    }
    if name.eq_ignore_ascii_case("IMCSC") {
        return apply_unary_complex(evaluator, node, numerics::csc);
    }
    if name.eq_ignore_ascii_case("IMCSCH") {
        return apply_unary_complex(evaluator, node, numerics::csch);
    }
    if name.eq_ignore_ascii_case("IMDIV") {
        return apply_binary_complex(evaluator, node, numerics::div);
    }
    if name.eq_ignore_ascii_case("IMEXP") {
        return apply_unary_complex(evaluator, node, numerics::exp);
    }
    if name.eq_ignore_ascii_case("IMLN") {
        return apply_unary_complex(evaluator, node, numerics::ln);
    }
    if name.eq_ignore_ascii_case("IMLOG10") {
        return apply_unary_complex(evaluator, node, |value| {
            let logarithm = numerics::ln(value)?;
            numerics::div(
                logarithm,
                Complex::checked(std::f64::consts::LN_10, 0.0, value.suffix)?,
            )
        });
    }
    if name.eq_ignore_ascii_case("IMLOG2") {
        return apply_unary_complex(evaluator, node, |value| {
            let logarithm = numerics::ln(value)?;
            numerics::div(
                logarithm,
                Complex::checked(std::f64::consts::LN_2, 0.0, value.suffix)?,
            )
        });
    }
    if name.eq_ignore_ascii_case("IMPOWER") {
        return apply_power(evaluator, node);
    }
    if name.eq_ignore_ascii_case("IMPRODUCT") {
        return apply_product(evaluator, node);
    }
    if name.eq_ignore_ascii_case("IMREAL") {
        return apply_unary_number(evaluator, node, |value| Ok(value.real));
    }
    if name.eq_ignore_ascii_case("IMSIN") {
        return apply_unary_complex(evaluator, node, numerics::sin);
    }
    if name.eq_ignore_ascii_case("IMSINH") {
        return apply_unary_complex(evaluator, node, numerics::sinh);
    }
    if name.eq_ignore_ascii_case("IMSEC") {
        return apply_unary_complex(evaluator, node, numerics::sec);
    }
    if name.eq_ignore_ascii_case("IMSECH") {
        return apply_unary_complex(evaluator, node, numerics::sech);
    }
    if name.eq_ignore_ascii_case("IMSQRT") {
        return apply_unary_complex(evaluator, node, numerics::sqrt);
    }
    if name.eq_ignore_ascii_case("IMSUB") {
        return apply_binary_complex(evaluator, node, complex_sub);
    }
    if name.eq_ignore_ascii_case("IMSUM") {
        return apply_sum(evaluator, node);
    }
    if name.eq_ignore_ascii_case("IMTAN") {
        return apply_unary_complex(evaluator, node, numerics::tan);
    }
    Err(EvaluationFailure::InvalidExpression(
        "unknown complex function reached evaluator",
    ))
}

fn pop_arguments<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    count: usize,
) -> EvaluationResult<ComplexArguments<'a>> {
    let mut arguments = ComplexArguments {
        values: Vec::new(),
        _reservation: None,
    };
    ensure_capacity(
        &mut arguments.values,
        &mut arguments._reservation,
        count,
        evaluator.limits.max_stack_entries,
        evaluator.context.execution,
        &evaluator.storage_budget,
        "formula complex function arguments",
    )?;
    let mut next_check = 0usize;
    for index in 0..count {
        check_scan(evaluator, index, &mut next_check)?;
        arguments.values.push(evaluator.pop_value()?);
    }
    arguments.values.reverse();
    Ok(arguments)
}

fn check_scan(
    evaluator: &mut Evaluator<'_, '_, '_>,
    index: usize,
    next_check: &mut usize,
) -> EvaluationResult<()> {
    if index >= *next_check {
        // A zero amount retains the caller's cancellation check without
        // charging the same argument again.  The enclosing fold charges one
        // unit per sequence element.
        evaluator.charge_work(0)?;
        *next_check = index.saturating_add(4096);
    }
    Ok(())
}

fn first_error(
    evaluator: &mut Evaluator<'_, '_, '_>,
    values: &[WorkingValue<'_>],
) -> EvaluationResult<Option<ScalarError>> {
    let mut next_check = 0usize;
    for (index, value) in values.iter().enumerate() {
        check_scan(evaluator, index, &mut next_check)?;
        if let WorkingValue::Error(error) = value {
            return Ok(Some(*error));
        }
    }
    Ok(None)
}

fn apply_complex<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
) -> EvaluationResult<()> {
    let count = node.child_count();
    if !(2..=3).contains(&count) {
        return evaluator.finish_invalid_arity(node);
    }
    let suffix_value = if count == 3 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let imaginary = evaluator.pop_value()?;
    let real = evaluator.pop_value()?;
    if let WorkingValue::Error(error) = &real {
        return evaluator.push_value(WorkingValue::Error(*error));
    }
    if let WorkingValue::Error(error) = &imaginary {
        return evaluator.push_value(WorkingValue::Error(*error));
    }
    if let Some(WorkingValue::Error(error)) = &suffix_value {
        return evaluator.push_value(WorkingValue::Error(*error));
    }
    let suffix = if let Some(value) = suffix_value {
        match value {
            WorkingValue::Text(text) => {
                evaluator.charge_bytes(text.text.len())?;
                if text.text == "i" {
                    'i'
                } else if text.text == "j" {
                    'j'
                } else {
                    return evaluator.push_value(WorkingValue::Error(ScalarError::Value));
                }
            },
            _ => return evaluator.push_value(WorkingValue::Error(ScalarError::Value)),
        }
    } else {
        'i'
    };
    let real = match super::to_number(real, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let imaginary = match super::to_number(imaginary, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    push_complex(evaluator, Complex::checked(real, imaginary, suffix))
}

fn apply_unary_complex<'a, F>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    operation: F,
) -> EvaluationResult<()>
where
    F: FnOnce(Complex) -> Result<Complex, ScalarError>,
{
    if node.child_count() != 1 {
        return evaluator.finish_invalid_arity(node);
    }
    let value = evaluator.pop_value()?;
    let value = match to_complex(value, evaluator)? {
        Ok(value) => operation(value),
        Err(error) => Err(error),
    };
    push_complex(evaluator, value)
}

fn apply_unary_number<'a, F>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    operation: F,
) -> EvaluationResult<()>
where
    F: FnOnce(Complex) -> Result<f64, ScalarError>,
{
    if node.child_count() != 1 {
        return evaluator.finish_invalid_arity(node);
    }
    let value = evaluator.pop_value()?;
    let value = match to_complex(value, evaluator)? {
        Ok(value) => operation(value),
        Err(error) => Err(error),
    };
    match value {
        Ok(value) if value.is_finite() => evaluator.push_value(WorkingValue::Number(value)),
        Ok(_) => evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
        Err(error) => evaluator.push_value(WorkingValue::Error(error)),
    }
}

fn apply_binary_complex<'a, F>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    operation: F,
) -> EvaluationResult<()>
where
    F: FnOnce(Complex, Complex) -> Result<Complex, ScalarError>,
{
    if node.child_count() != 2 {
        return evaluator.finish_invalid_arity(node);
    }
    let right = evaluator.pop_value()?;
    let left = evaluator.pop_value()?;
    if let WorkingValue::Error(error) = &left {
        return evaluator.push_value(WorkingValue::Error(*error));
    }
    if let WorkingValue::Error(error) = &right {
        return evaluator.push_value(WorkingValue::Error(*error));
    }
    let left = match to_complex(left, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let right = match to_complex(right, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    push_complex(evaluator, operation(left, right))
}

fn apply_power<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    if node.child_count() != 2 {
        return evaluator.finish_invalid_arity(node);
    }
    let right = evaluator.pop_value()?;
    let left = evaluator.pop_value()?;
    if let WorkingValue::Error(error) = &left {
        return evaluator.push_value(WorkingValue::Error(*error));
    }
    if let WorkingValue::Error(error) = &right {
        return evaluator.push_value(WorkingValue::Error(*error));
    }
    let left = match to_complex(left, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let right = match to_complex(right, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    if left.is_zero() {
        return evaluator.push_value(WorkingValue::Error(ScalarError::Number));
    }
    let logarithm = match numerics::ln(left) {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let product = match complex_mul(logarithm, right) {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    push_complex(evaluator, numerics::exp(product))
}

fn apply_product<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
) -> EvaluationResult<()> {
    let count = node.child_count();
    if count == 0 {
        return evaluator.finish_invalid_arity(node);
    }
    let values = pop_arguments(evaluator, count)?;
    if let Some(error) = first_error(evaluator, values.as_slice())? {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let mut result = Complex::one('i');
    let mut generated_error = None;
    for value in values {
        evaluator.charge_work(1)?;
        let value = match to_complex(value, evaluator)? {
            Ok(value) => value,
            Err(error) => {
                if generated_error.is_none() {
                    generated_error = Some(error);
                }
                continue;
            },
        };
        result.suffix = choose_suffix(result.suffix, value.suffix);
        if generated_error.is_none() {
            match complex_mul(result, value) {
                Ok(value) => result = value,
                Err(error) => generated_error = Some(error),
            }
        }
    }
    match generated_error {
        Some(error) => evaluator.push_value(WorkingValue::Error(error)),
        None => evaluator.push_value(WorkingValue::Complex(result)),
    }
}

fn apply_sum<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    let count = node.child_count();
    if count == 0 {
        return evaluator.push_value(WorkingValue::Number(0.0));
    }
    let values = pop_arguments(evaluator, count)?;
    if let Some(error) = first_error(evaluator, values.as_slice())? {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let mut result = Complex::zero('i');
    let mut generated_error = None;
    for value in values {
        evaluator.charge_work(1)?;
        let value = match value {
            WorkingValue::Text(text) => {
                evaluator.charge_bytes(text.text.len())?;
                match parse_complex_text(text.text.as_ref()) {
                    Ok(value) => value,
                    Err(_) => continue,
                }
            },
            value => match to_complex(value, evaluator)? {
                Ok(value) => value,
                Err(error) => {
                    if generated_error.is_none() {
                        generated_error = Some(error);
                    }
                    continue;
                },
            },
        };
        result.suffix = choose_suffix(result.suffix, value.suffix);
        if generated_error.is_none() {
            match complex_add(result, value) {
                Ok(value) => result = value,
                Err(error) => generated_error = Some(error),
            }
        }
    }
    match generated_error {
        Some(error) => evaluator.push_value(WorkingValue::Error(error)),
        None => evaluator.push_value(WorkingValue::Complex(result)),
    }
}

fn push_complex<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: Result<Complex, ScalarError>,
) -> EvaluationResult<()> {
    match value {
        Ok(value) => evaluator.push_value(WorkingValue::Complex(value)),
        Err(error) => evaluator.push_value(WorkingValue::Error(error)),
    }
}

pub(super) fn to_complex<'a>(
    value: WorkingValue<'a>,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<Complex, ScalarError>> {
    match value {
        WorkingValue::Number(value) => Ok(Complex::checked(value, 0.0, 'i')),
        WorkingValue::Logical(value) => {
            Ok(Complex::checked(if value { 1.0 } else { 0.0 }, 0.0, 'i'))
        },
        WorkingValue::Complex(value) => Ok(Ok(value)),
        WorkingValue::Text(text) => {
            evaluator.charge_bytes(text.text.len())?;
            Ok(parse_complex_text(text.text.as_ref()))
        },
        WorkingValue::Error(error) => Ok(Err(error)),
    }
}

fn parse_complex_text(text: &str) -> Result<Complex, ScalarError> {
    if text.is_empty() {
        return Err(ScalarError::Value);
    }
    let (body, suffix) = match text.as_bytes().last().copied() {
        Some(b'i') => (&text[..text.len() - 1], 'i'),
        Some(b'j') => (&text[..text.len() - 1], 'j'),
        _ => {
            let value = parse_finite_number(text)?;
            return Complex::checked(value, 0.0, 'i');
        },
    };

    if body.is_empty() || body == "+" || body == "-" {
        return Complex::checked(0.0, if body == "-" { -1.0 } else { 1.0 }, suffix);
    }

    let Some(first_end) = number_prefix(body) else {
        return Err(ScalarError::Value);
    };
    if first_end == body.len() {
        let imaginary = parse_finite_number(body)?;
        return Complex::checked(0.0, imaginary, suffix);
    }
    let separator = body.as_bytes().get(first_end).copied();
    if !matches!(separator, Some(b'+' | b'-')) {
        return Err(ScalarError::Value);
    }
    let real = parse_finite_number(&body[..first_end])?;
    let imaginary_text = &body[first_end..];
    let imaginary = if imaginary_text == "+" {
        1.0
    } else if imaginary_text == "-" {
        -1.0
    } else {
        parse_finite_number(imaginary_text)?
    };
    Complex::checked(real, imaginary, suffix)
}

fn parse_finite_number(text: &str) -> Result<f64, ScalarError> {
    if text.is_empty() {
        return Err(ScalarError::Value);
    }
    let value = fast_float2::parse::<f64, _>(text).map_err(|_| ScalarError::Value)?;
    if value.is_finite() {
        Ok(value)
    } else {
        Err(ScalarError::Number)
    }
}

fn number_prefix(text: &str) -> Option<usize> {
    let bytes = text.as_bytes();
    let mut index = 0usize;
    if matches!(bytes.first(), Some(b'+' | b'-')) {
        index += 1;
    }
    let mut digits = 0usize;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
        digits += 1;
    }
    if bytes.get(index) == Some(&b'.') {
        index += 1;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
            digits += 1;
        }
    }
    if digits == 0 {
        return None;
    }
    if matches!(bytes.get(index), Some(b'e' | b'E')) {
        index += 1;
        if matches!(bytes.get(index), Some(b'+' | b'-')) {
            index += 1;
        }
        let exponent_start = index;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
        }
        if index == exponent_start {
            return None;
        }
    }
    Some(index)
}

fn choose_suffix(left: char, right: char) -> char {
    if left == 'j' || right == 'j' {
        'j'
    } else {
        'i'
    }
}

fn unary_abs(value: Complex) -> Result<f64, ScalarError> {
    let value = value.real.hypot(value.imaginary);
    value
        .is_finite()
        .then_some(value)
        .ok_or(ScalarError::Number)
}

fn complex_argument(value: Complex) -> f64 {
    if value.is_zero() {
        return 0.0;
    }
    let mut angle = value.imaginary.atan2(value.real);
    if angle == -std::f64::consts::PI {
        angle = std::f64::consts::PI;
    } else if angle == 0.0 {
        angle = 0.0;
    }
    angle
}

pub(super) fn complex_add(left: Complex, right: Complex) -> Result<Complex, ScalarError> {
    Complex::checked(
        left.real + right.real,
        left.imaginary + right.imaginary,
        choose_suffix(left.suffix, right.suffix),
    )
}

fn complex_sub(left: Complex, right: Complex) -> Result<Complex, ScalarError> {
    Complex::checked(
        left.real - right.real,
        left.imaginary - right.imaginary,
        choose_suffix(left.suffix, right.suffix),
    )
}

pub(super) fn complex_mul(left: Complex, right: Complex) -> Result<Complex, ScalarError> {
    numerics::mul(left, right)
}

/// Convert a complex value to formula text for the ordinary concatenation
/// operator.  The bounded stack formatter avoids an intermediate infallible
/// `String` allocation; the final owned text follows the evaluator's normal
/// reservation contract.
pub(super) fn to_text<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    value: Complex,
) -> EvaluationResult<TextValue<'a>> {
    let mut number = NumberText::new();
    write!(&mut number, "{value}")
        .map_err(|_| EvaluationFailure::InvalidExpression("complex-to-text formatting failed"))?;
    let rendered = number
        .as_str()
        .map_err(|_| EvaluationFailure::InvalidExpression("invalid complex text"))?;
    evaluator.charge_bytes(rendered.len())?;
    if rendered.len() > evaluator.limits.max_text_bytes {
        return Err(local_limit(
            Resource::Memory,
            u64::try_from(rendered.len()).unwrap_or(u64::MAX),
            u64::try_from(evaluator.limits.max_text_bytes).unwrap_or(u64::MAX),
        ));
    }
    let reservation = evaluator.reserve_storage(rendered.len(), "formula complex text")?;
    let mut output = String::new();
    if !rendered.is_empty() {
        output.try_reserve_exact(rendered.len()).map_err(|source| {
            EvaluationFailure::Allocation {
                resource: "formula complex text",
                source,
            }
        })?;
    }
    output.push_str(rendered);
    Ok(TextValue::owned(output, reservation))
}

#[cfg(test)]
mod tests {
    use super::{number_prefix, parse_complex_text};

    #[test]
    fn parses_real_and_imaginary_forms() {
        let value = parse_complex_text("3+4i").expect("complex");
        assert_eq!(value.real(), 3.0);
        assert_eq!(value.imaginary(), 4.0);
        assert_eq!(value.suffix(), 'i');
        let value = parse_complex_text("-2j").expect("imaginary");
        assert_eq!(value.real(), 0.0);
        assert_eq!(value.imaginary(), -2.0);
        assert_eq!(value.suffix(), 'j');
    }

    #[test]
    fn exponent_sign_is_not_a_complex_separator() {
        assert_eq!(number_prefix("3e-2"), Some(4));
        let value = parse_complex_text("3e-2i").expect("imaginary exponent");
        assert_eq!(value.real(), 0.0);
        assert!((value.imaginary() - 0.03).abs() < 1e-15);
    }
}
