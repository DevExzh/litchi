//! OpenFormula 1.4 section 6.19 radix conversion functions.
//!
//! The implementation deliberately keeps the integer machinery private to
//! the evaluator.  `BASE` and `DECIMAL` use a fixed 1024-bit magnitude, which
//! is enough for every finite integer `f64`; the direct `BIN`/`OCT`/`HEX`
//! conversions use their specified 10-, 30-, and 40-bit signed widths.

use super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, NumberText, ScalarError, TextValue,
    WorkingValue, local_limit, map_execution_error,
};
use litchi_core::{Reservation, Resource};
use std::fmt::Write as FmtWrite;

const MAX_LIMBS: usize = 32;
const MAX_RADIX_DIGITS: usize = 1024;
const MAX_DIRECT_DIGITS: usize = 10;
const BITS_PER_LIMB: usize = 32;
const DIGITS: &[u8; 36] = b"0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

#[derive(Clone, Copy)]
enum Radix {
    Binary,
    Octal,
    Hexadecimal,
}

impl Radix {
    const fn value(self) -> u32 {
        match self {
            Self::Binary => 2,
            Self::Octal => 8,
            Self::Hexadecimal => 16,
        }
    }

    const fn width(self) -> u32 {
        match self {
            Self::Binary => 10,
            Self::Octal => 30,
            Self::Hexadecimal => 40,
        }
    }

    const fn output_digits(self) -> usize {
        MAX_DIRECT_DIGITS
    }
}

#[derive(Clone, Copy)]
enum Function {
    Base,
    Decimal,
    FromDecimal(Radix),
    Between(Radix, Radix),
    ToDecimal(Radix),
}

fn function(name: &str) -> Option<Function> {
    if name.eq_ignore_ascii_case("BASE") {
        return Some(Function::Base);
    }
    if name.eq_ignore_ascii_case("DECIMAL") {
        return Some(Function::Decimal);
    }
    if name.eq_ignore_ascii_case("BIN2DEC") {
        return Some(Function::ToDecimal(Radix::Binary));
    }
    if name.eq_ignore_ascii_case("BIN2HEX") {
        return Some(Function::Between(Radix::Binary, Radix::Hexadecimal));
    }
    if name.eq_ignore_ascii_case("BIN2OCT") {
        return Some(Function::Between(Radix::Binary, Radix::Octal));
    }
    if name.eq_ignore_ascii_case("DEC2BIN") {
        return Some(Function::FromDecimal(Radix::Binary));
    }
    if name.eq_ignore_ascii_case("DEC2HEX") {
        return Some(Function::FromDecimal(Radix::Hexadecimal));
    }
    if name.eq_ignore_ascii_case("DEC2OCT") {
        return Some(Function::FromDecimal(Radix::Octal));
    }
    if name.eq_ignore_ascii_case("HEX2BIN") {
        return Some(Function::Between(Radix::Hexadecimal, Radix::Binary));
    }
    if name.eq_ignore_ascii_case("HEX2DEC") {
        return Some(Function::ToDecimal(Radix::Hexadecimal));
    }
    if name.eq_ignore_ascii_case("HEX2OCT") {
        return Some(Function::Between(Radix::Hexadecimal, Radix::Octal));
    }
    if name.eq_ignore_ascii_case("OCT2BIN") {
        return Some(Function::Between(Radix::Octal, Radix::Binary));
    }
    if name.eq_ignore_ascii_case("OCT2DEC") {
        return Some(Function::ToDecimal(Radix::Octal));
    }
    if name.eq_ignore_ascii_case("OCT2HEX") {
        return Some(Function::Between(Radix::Octal, Radix::Hexadecimal));
    }
    None
}

pub(super) fn is_radix_function(name: &str) -> bool {
    function(name).is_some()
}

/// Apply a section 6.19 function after all of its arguments were evaluated.
/// This handler is intentionally outlined so adding the uncommon conversion
/// family does not enlarge the evaluator's scalar hot loop.
#[inline(never)]
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    match function(name) {
        Some(Function::Base) => apply_base(evaluator, node),
        Some(Function::Decimal) => apply_decimal(evaluator, node),
        Some(
            function
            @ (Function::FromDecimal(_) | Function::Between(_, _) | Function::ToDecimal(_)),
        ) => apply_direct(evaluator, node, function),
        None => Err(EvaluationFailure::InvalidExpression(
            "unknown radix function reached evaluator",
        )),
    }
}

#[derive(Clone, Copy)]
struct Uint1024 {
    limbs: [u32; MAX_LIMBS],
}

impl Uint1024 {
    const fn zero() -> Self {
        Self {
            limbs: [0; MAX_LIMBS],
        }
    }

    const fn from_u64(value: u64) -> Self {
        let mut limbs = [0; MAX_LIMBS];
        limbs[0] = value as u32;
        limbs[1] = (value >> 32) as u32;
        Self { limbs }
    }

    fn from_f64_integer(value: f64) -> Result<Self, ScalarError> {
        if !value.is_finite() || value < 0.0 {
            return Err(ScalarError::Number);
        }
        let value = value.trunc();
        if value == 0.0 {
            return Ok(Self::zero());
        }

        let bits = value.to_bits();
        let exponent = ((bits >> 52) & 0x7ff) as i32;
        if exponent == 0 {
            return Ok(Self::zero());
        }
        if exponent == 0x7ff {
            return Err(ScalarError::Number);
        }
        let significand = (1_u64 << 52) | (bits & ((1_u64 << 52) - 1));
        let shift = exponent - 1023 - 52;
        if shift < 0 {
            let right = usize::try_from(-shift).map_err(|_| ScalarError::Number)?;
            return Ok(Self::from_u64(significand >> right));
        }

        let shift = usize::try_from(shift).map_err(|_| ScalarError::Number)?;
        let highest = shift.checked_add(52).ok_or(ScalarError::Number)?;
        if highest >= MAX_LIMBS * BITS_PER_LIMB {
            return Err(ScalarError::Number);
        }
        let mut result = Self::zero();
        for bit in 0..=52 {
            if (significand & (1_u64 << bit)) != 0 {
                result.set_bit(shift + bit);
            }
        }
        Ok(result)
    }

    fn set_bit(&mut self, index: usize) {
        self.limbs[index / BITS_PER_LIMB] |= 1_u32 << (index % BITS_PER_LIMB);
    }

    fn is_zero(self) -> bool {
        self.limbs.iter().all(|limb| *limb == 0)
    }

    fn highest_bit(self) -> Option<usize> {
        for (index, limb) in self.limbs.iter().enumerate().rev() {
            if *limb != 0 {
                return Some(index * BITS_PER_LIMB + (31 - limb.leading_zeros() as usize));
            }
        }
        None
    }

    fn bit(self, index: usize) -> bool {
        index < MAX_LIMBS * BITS_PER_LIMB
            && (self.limbs[index / BITS_PER_LIMB] & (1_u32 << (index % BITS_PER_LIMB))) != 0
    }

    fn any_below(self, bits: usize) -> bool {
        let full = bits / BITS_PER_LIMB;
        if self.limbs[..full.min(MAX_LIMBS)]
            .iter()
            .any(|limb| *limb != 0)
        {
            return true;
        }
        if full >= MAX_LIMBS {
            return false;
        }
        let remainder = bits % BITS_PER_LIMB;
        remainder != 0 && (self.limbs[full] & ((1_u32 << remainder) - 1)) != 0
    }

    fn to_u64(self) -> Option<u64> {
        if self.limbs[2..].iter().any(|limb| *limb != 0) {
            return None;
        }
        Some((self.limbs[1] as u64) << 32 | self.limbs[0] as u64)
    }

    /// Multiply by a small radix and add a digit.  `true` means that a bit
    /// escaped the fixed 1024-bit domain.
    fn mul_small_add(&mut self, radix: u32, digit: u32) -> bool {
        let mut carry = u64::from(digit);
        for limb in &mut self.limbs {
            let value = u64::from(*limb) * u64::from(radix) + carry;
            *limb = value as u32;
            carry = value >> 32;
        }
        carry != 0
    }

    fn div_small(&mut self, radix: u32) -> u32 {
        let mut remainder = 0_u64;
        for limb in self.limbs.iter_mut().rev() {
            let value = (remainder << 32) | u64::from(*limb);
            *limb = (value / u64::from(radix)) as u32;
            remainder = value % u64::from(radix);
        }
        remainder as u32
    }

    /// Convert an arbitrary non-negative integer to finite `f64` with one
    /// IEEE nearest-even rounding step.
    fn to_f64_nearest_even(self) -> Option<f64> {
        let highest = match self.highest_bit() {
            Some(highest) => highest,
            None => return Some(0.0),
        };
        if highest <= 52 {
            let mut result = 0.0_f64;
            for bit in (0..=highest).rev() {
                result = result * 2.0 + if self.bit(bit) { 1.0 } else { 0.0 };
            }
            return Some(result);
        }

        let shift = highest - 52;
        let mut significand = 0_u64;
        for offset in 0..=52 {
            if self.bit(highest - offset) {
                significand |= 1_u64 << (52 - offset);
            }
        }
        let round = self.bit(shift - 1);
        let sticky = self.any_below(shift - 1);
        if round && (sticky || (significand & 1) != 0) {
            significand = significand.checked_add(1)?;
        }

        let mut exponent = highest;
        if significand == 1_u64 << 53 {
            significand >>= 1;
            exponent = exponent.checked_add(1)?;
        }
        if exponent > 1023 {
            return None;
        }
        let exponent_bits = u64::try_from(exponent + 1023).ok()?;
        let fraction = significand & ((1_u64 << 52) - 1);
        Some(f64::from_bits((exponent_bits << 52) | fraction))
    }
}

#[derive(Clone, Copy)]
struct SignedInteger {
    negative: bool,
    magnitude: Uint1024,
}

impl SignedInteger {
    const fn zero() -> Self {
        Self {
            negative: false,
            magnitude: Uint1024::zero(),
        }
    }
}

fn formula_error(value: &WorkingValue<'_>) -> Option<ScalarError> {
    match value {
        WorkingValue::Error(error) => Some(*error),
        _ => None,
    }
}

fn apply_base<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    let count = node.child_count();
    if !(2..=3).contains(&count) {
        return evaluator.finish_invalid_arity(node);
    }

    let minimum = if count == 3 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let radix = evaluator.pop_value()?;
    let value = evaluator.pop_value()?;

    if let Some(error) = formula_error(&value) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(error) = formula_error(&radix) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(minimum) = &minimum {
        if let Some(error) = formula_error(minimum) {
            return evaluator.push_value(WorkingValue::Error(error));
        }
    }

    let value = match super::to_integer(value, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let radix = match super::to_integer(radix, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    if value < 0.0 || !value.is_finite() || !(2.0..=36.0).contains(&radix) {
        return evaluator.push_value(WorkingValue::Error(ScalarError::Number));
    }

    let minimum = match minimum {
        Some(value) => match super::to_integer(value, evaluator)? {
            Ok(value) if value.is_finite() && (0.0..).contains(&value) => {
                bounded_length(value, evaluator.limits.max_text_bytes)?
            },
            Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
            Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
        },
        None => 0,
    };
    let radix = radix as u32;
    let value = match Uint1024::from_f64_integer(value) {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    render_unsigned(evaluator, value, radix, minimum, None)
}

fn apply_decimal<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
) -> EvaluationResult<()> {
    if node.child_count() != 2 {
        return evaluator.finish_invalid_arity(node);
    }
    let radix = evaluator.pop_value()?;
    let value = evaluator.pop_value()?;
    if let Some(error) = formula_error(&value) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(error) = formula_error(&radix) {
        return evaluator.push_value(WorkingValue::Error(error));
    }

    let radix = match super::to_integer(radix, evaluator)? {
        Ok(value) if value.is_finite() && (2.0..=36.0).contains(&value) => value as u32,
        Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let text = super::to_text(value, evaluator)?;
    let parsed = parse_decimal_text(evaluator, text.text.as_ref(), radix)?;
    drop(text);
    let value = match parsed {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let value = match value.to_f64_nearest_even() {
        Some(value) if value.is_finite() => WorkingValue::Number(value),
        _ => WorkingValue::Error(ScalarError::Number),
    };
    evaluator.push_value(value)
}

fn apply_direct<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    function: Function,
) -> EvaluationResult<()> {
    let to_decimal = matches!(function, Function::ToDecimal(_));
    let count = node.child_count();
    if (to_decimal && count != 1) || (!to_decimal && !(1..=2).contains(&count)) {
        return evaluator.finish_invalid_arity(node);
    }

    let digits = if to_decimal && count == 1 {
        None
    } else if count == 2 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let value = evaluator.pop_value()?;
    if let Some(error) = formula_error(&value) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(digits) = &digits {
        if let Some(error) = formula_error(digits) {
            return evaluator.push_value(WorkingValue::Error(error));
        }
    }

    match function {
        Function::ToDecimal(source) => {
            let value = match parse_fixed_value(evaluator, value, source)? {
                Ok(value) => value,
                Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
            };
            let number = signed_to_number(value);
            evaluator.push_value(WorkingValue::Number(number))
        },
        Function::FromDecimal(target) => {
            let value = match parse_decimal_signed_value(evaluator, value)? {
                Ok(value) => value,
                Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
            };
            render_signed(evaluator, value, target, digits)
        },
        Function::Between(source, target) => {
            let value = match parse_fixed_value(evaluator, value, source)? {
                Ok(value) => value,
                Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
            };
            render_signed(evaluator, value, target, digits)
        },
        Function::Base | Function::Decimal => Err(EvaluationFailure::InvalidExpression(
            "non-direct radix function reached direct handler",
        )),
    }
}

fn parse_decimal_signed_value<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<SignedInteger, ScalarError>> {
    match value {
        WorkingValue::Text(text) => parse_signed_decimal_text(evaluator, text.text.as_ref()),
        WorkingValue::Number(value) => {
            if !value.is_finite() {
                return Ok(Err(ScalarError::Number));
            }
            if value.fract() != 0.0 {
                return Ok(Err(ScalarError::Value));
            }
            let value = value.trunc();
            let negative = value < 0.0;
            let magnitude = Uint1024::from_f64_integer(value.abs())
                .map_err(|_| EvaluationFailure::InvalidExpression("decimal integer overflow"))?;
            Ok(Ok(if magnitude.is_zero() {
                SignedInteger::zero()
            } else {
                SignedInteger {
                    negative,
                    magnitude,
                }
            }))
        },
        WorkingValue::Logical(value) => Ok(Ok(SignedInteger {
            negative: false,
            magnitude: Uint1024::from_u64(if value { 1 } else { 0 }),
        })),
        WorkingValue::Complex(_) => Ok(Err(ScalarError::Value)),
        WorkingValue::Error(error) => Ok(Err(error)),
    }
}

fn parse_signed_decimal_text(
    evaluator: &mut Evaluator<'_, '_, '_>,
    text: &str,
) -> EvaluationResult<Result<SignedInteger, ScalarError>> {
    evaluator.charge_bytes(text.len())?;
    let (negative, digits) = match text.strip_prefix('-') {
        Some(digits) => (true, digits),
        None => (false, text),
    };
    if digits.is_empty() {
        return Ok(Err(ScalarError::Value));
    }
    let mut value = Uint1024::zero();
    let mut overflow = false;
    let mut next_check = 4096usize;
    for (index, byte) in digits.bytes().enumerate() {
        if index >= next_check {
            evaluator
                .context
                .execution
                .check()
                .map_err(map_execution_error)?;
            next_check = index.saturating_add(4096);
        }
        let digit = match byte.checked_sub(b'0') {
            Some(digit @ 0..=9) => u32::from(digit),
            _ => return Ok(Err(ScalarError::Value)),
        };
        overflow |= value.mul_small_add(10, digit);
    }
    if overflow {
        return Ok(Err(ScalarError::Number));
    }
    Ok(Ok(if value.is_zero() {
        SignedInteger::zero()
    } else {
        SignedInteger {
            negative,
            magnitude: value,
        }
    }))
}

fn parse_fixed_value<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    value: WorkingValue<'a>,
    radix: Radix,
) -> EvaluationResult<Result<SignedInteger, ScalarError>> {
    match value {
        WorkingValue::Text(text) => parse_fixed_text(evaluator, text.text.as_ref(), radix),
        WorkingValue::Number(value) => {
            if !value.is_finite() {
                return Ok(Err(ScalarError::Number));
            }
            if value.fract() != 0.0 {
                return Ok(Err(ScalarError::Value));
            }
            let value = if value == 0.0 { 0.0 } else { value };
            let mut number = NumberText::new();
            write!(&mut number, "{value}")
                .map_err(|_| EvaluationFailure::InvalidExpression("number text overflow"))?;
            let text = number
                .as_str()
                .map_err(|_| EvaluationFailure::InvalidExpression("number text is not UTF-8"))?;
            evaluator.charge_bytes(text.len())?;
            parse_fixed_text(evaluator, text, radix)
        },
        WorkingValue::Logical(value) => {
            parse_fixed_text(evaluator, if value { "1" } else { "0" }, radix)
        },
        WorkingValue::Complex(_) => Ok(Err(ScalarError::Value)),
        WorkingValue::Error(error) => Ok(Err(error)),
    }
}

fn parse_fixed_text(
    evaluator: &mut Evaluator<'_, '_, '_>,
    text: &str,
    radix: Radix,
) -> EvaluationResult<Result<SignedInteger, ScalarError>> {
    evaluator.charge_bytes(text.len())?;
    if text.is_empty() {
        return Ok(Err(ScalarError::Value));
    }
    let bytes = text.as_bytes();
    let mut raw = 0_u64;
    for (index, byte) in bytes.iter().copied().enumerate() {
        if index % 4096 == 0 {
            evaluator
                .context
                .execution
                .check()
                .map_err(map_execution_error)?;
        }
        let digit = match radix_digit(byte) {
            Some(digit) if digit < radix.value() => u64::from(digit),
            _ => return Ok(Err(ScalarError::Value)),
        };
        if index >= MAX_DIRECT_DIGITS {
            return Ok(Err(ScalarError::Number));
        }
        raw = raw * u64::from(radix.value()) + digit;
    }
    let width = radix.width();
    let sign = 1_u64 << (width - 1);
    let modulus = 1_u64 << width;
    if raw & sign == 0 {
        Ok(Ok(SignedInteger {
            negative: false,
            magnitude: Uint1024::from_u64(raw),
        }))
    } else {
        Ok(Ok(SignedInteger {
            negative: true,
            magnitude: Uint1024::from_u64(modulus - raw),
        }))
    }
}

fn parse_decimal_text(
    evaluator: &mut Evaluator<'_, '_, '_>,
    text: &str,
    radix: u32,
) -> EvaluationResult<Result<Uint1024, ScalarError>> {
    evaluator.charge_bytes(text.len())?;
    let bytes = text.as_bytes();
    let mut start = 0usize;
    let mut next_check = 4096usize;
    while start < bytes.len() && matches!(bytes[start], b' ' | b'\t') {
        if start >= next_check {
            evaluator
                .context
                .execution
                .check()
                .map_err(map_execution_error)?;
            next_check = start.saturating_add(4096);
        }
        start += 1;
    }
    let mut end = bytes.len();
    if radix == 16 {
        if bytes
            .get(start..start.saturating_add(2))
            .is_some_and(|prefix| prefix[0] == b'0' && matches!(prefix[1], b'x' | b'X'))
        {
            start += 2;
        } else if bytes
            .get(start)
            .is_some_and(|byte| matches!(byte, b'x' | b'X'))
        {
            start += 1;
        }
        if bytes
            .get(end.wrapping_sub(1))
            .is_some_and(|byte| matches!(byte, b'h' | b'H'))
        {
            end = end.saturating_sub(1);
        }
    } else if radix == 2
        && bytes
            .get(end.wrapping_sub(1))
            .is_some_and(|byte| matches!(byte, b'b' | b'B'))
    {
        end = end.saturating_sub(1);
    }
    if start >= end {
        return Ok(Err(ScalarError::Value));
    }

    let mut value = Uint1024::zero();
    let mut overflow = false;
    next_check = 4096;
    for (index, byte) in bytes[start..end].iter().copied().enumerate() {
        if index >= next_check {
            evaluator
                .context
                .execution
                .check()
                .map_err(map_execution_error)?;
            next_check = index.saturating_add(4096);
        }
        let digit = match radix_digit(byte) {
            Some(digit) if digit < radix => digit,
            _ => return Ok(Err(ScalarError::Value)),
        };
        overflow |= value.mul_small_add(radix, digit);
    }
    if overflow {
        return Ok(Err(ScalarError::Number));
    }
    Ok(Ok(value))
}

fn radix_digit(byte: u8) -> Option<u32> {
    match byte {
        b'0'..=b'9' => Some(u32::from(byte - b'0')),
        b'A'..=b'Z' => Some(u32::from(byte - b'A') + 10),
        b'a'..=b'z' => Some(u32::from(byte - b'a') + 10),
        _ => None,
    }
}

fn bounded_length(value: f64, maximum: usize) -> EvaluationResult<usize> {
    if value > maximum as f64 || value >= usize::MAX as f64 {
        return Err(local_limit(
            Resource::Memory,
            observed_f64(value),
            u64::try_from(maximum).unwrap_or(u64::MAX),
        ));
    }
    let value = value as usize;
    if value > maximum {
        return Err(local_limit(
            Resource::Memory,
            u64::try_from(value).unwrap_or(u64::MAX),
            u64::try_from(maximum).unwrap_or(u64::MAX),
        ));
    }
    Ok(value)
}

fn observed_f64(value: f64) -> u64 {
    if !value.is_finite() || value >= u64::MAX as f64 {
        u64::MAX
    } else if value <= 0.0 {
        0
    } else {
        value as u64
    }
}

fn signed_to_number(value: SignedInteger) -> f64 {
    let magnitude = value.magnitude.to_u64().unwrap_or(u64::MAX) as f64;
    if value.negative {
        -magnitude
    } else {
        magnitude
    }
}

fn signed_to_raw(value: SignedInteger, width: u32) -> Result<(u64, bool), ScalarError> {
    let magnitude = value.magnitude.to_u64().ok_or(ScalarError::Number)?;
    let sign_limit = 1_u64 << (width - 1);
    if value.negative {
        if magnitude > sign_limit {
            return Err(ScalarError::Number);
        }
        let raw = if magnitude == 0 {
            0
        } else {
            (1_u64 << width) - magnitude
        };
        Ok((raw, magnitude != 0))
    } else {
        if magnitude >= sign_limit {
            return Err(ScalarError::Number);
        }
        Ok((magnitude, false))
    }
}

fn render_signed<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: SignedInteger,
    target: Radix,
    digits: Option<WorkingValue<'a>>,
) -> EvaluationResult<()> {
    let (raw, negative) = match signed_to_raw(value, target.width()) {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };

    // A negative result has a fixed sign-extension width.  Its Digits value
    // is intentionally not converted, although an already evaluated formula
    // error in that argument was propagated by the caller.
    let minimum = if negative {
        target.output_digits()
    } else {
        match digits {
            Some(value) => match super::to_integer(value, evaluator)? {
                Ok(value) if (0.0..=10.0).contains(&value) => value as usize,
                Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
                Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
            },
            None => 0,
        }
    };

    render_unsigned(
        evaluator,
        Uint1024::from_u64(raw),
        target.value(),
        minimum,
        negative.then_some(target.output_digits()),
    )
}

fn render_unsigned<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    mut value: Uint1024,
    radix: u32,
    minimum: usize,
    fixed: Option<usize>,
) -> EvaluationResult<()> {
    let mut reversed = [0_u8; MAX_RADIX_DIGITS];
    let mut length = 0usize;
    if value.is_zero() {
        reversed[0] = 0;
        length = 1;
    } else {
        while !value.is_zero() {
            if length == reversed.len() {
                return evaluator.push_value(WorkingValue::Error(ScalarError::Number));
            }
            evaluator.charge_work(1)?;
            reversed[length] = value.div_small(radix) as u8;
            length += 1;
        }
    }
    let output_length = fixed.unwrap_or(minimum).max(length);
    let (mut output, reservation) = reserve_output(evaluator, output_length)?;
    let mut next_check = 4096usize;
    for index in 0..output_length {
        if index >= next_check {
            evaluator
                .context
                .execution
                .check()
                .map_err(map_execution_error)?;
            next_check = index.saturating_add(4096);
        }
        let digit = if index < output_length - length {
            0
        } else {
            reversed[length - 1 - (index - (output_length - length))]
        };
        output.push(char::from(DIGITS[usize::from(digit)]));
    }
    evaluator.push_value(WorkingValue::Text(TextValue::owned(output, reservation)))
}

fn reserve_output(
    evaluator: &mut Evaluator<'_, '_, '_>,
    length: usize,
) -> EvaluationResult<(String, Reservation)> {
    if length > evaluator.limits.max_text_bytes {
        return Err(local_limit(
            Resource::Memory,
            u64::try_from(length).unwrap_or(u64::MAX),
            u64::try_from(evaluator.limits.max_text_bytes).unwrap_or(u64::MAX),
        ));
    }
    evaluator.charge_bytes(length)?;
    let reservation = evaluator.reserve_storage(length, "formula radix text")?;
    let mut output = String::new();
    if length != 0 {
        output
            .try_reserve_exact(length)
            .map_err(|source| EvaluationFailure::Allocation {
                resource: "formula radix text",
                source,
            })?;
    }
    Ok((output, reservation))
}
