//! OpenFormula 1.4 section 6.16 discrete numeric functions.
//!
//! The functions in this module all operate on finite binary64 values.
//! COMBIN and COMBINA explicitly apply `INT` before their constraint checks.
//! This profile uses `INT` (floor toward negative infinity) for the other
//! integer-labelled functions after scalar Number conversion, except for the
//! explicitly strict LCM constraint. Integer values are retained in a small
//! fixed-width unsigned
//! representation while the mathematical operation is performed.  This is
//! important for two reasons: a represented integer above 2^53 is still a
//! useful input to GCD/LCM, and adding `2^53;1;1` must retain both increments
//! when it is used as a multinomial total.
//!
//! The fixed-width state is large enough for every finite binary64 integer
//! operand (1024 bits).  A result is converted once with round-to-nearest-even
//! at the Number boundary.  A result whose exact integer cannot be converted
//! to a finite binary64 value is `#NUM!`; an exact value just beyond the
//! largest finite value may round back to that value.  Combinations use an
//! exact divide-before-multiply recurrence.  The smaller side of a binomial
//! is selected first, and a smaller side of 1024 or more proves that the
//! result is outside the finite Number range, so no unbounded loop is needed.
//!
//! [`DiscreteFold`] is shared with the resolver-backed value
//! evaluator.  It accepts already-admitted finite Numbers one at a time and
//! keeps formula errors separate from generated conversion/domain errors.  It
//! does not retain a reference or allocate a staging vector.

use super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, UnsupportedKind,
    WorkingValue,
};

const BITS_PER_LIMB: usize = 32;
const MAX_LIMBS: usize = 32;
const MAX_BITS: usize = BITS_PER_LIMB * MAX_LIMBS;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Function {
    Combin,
    Combina,
    Delta,
    Even,
    Fact,
    FactDouble,
    Gcd,
    GEstep,
    Lcm,
    Multinomial,
    Odd,
}

fn function(name: &str) -> Option<Function> {
    Some(if name.eq_ignore_ascii_case("COMBIN") {
        Function::Combin
    } else if name.eq_ignore_ascii_case("COMBINA") {
        Function::Combina
    } else if name.eq_ignore_ascii_case("DELTA") {
        Function::Delta
    } else if name.eq_ignore_ascii_case("EVEN") {
        Function::Even
    } else if name.eq_ignore_ascii_case("FACT") {
        Function::Fact
    } else if name.eq_ignore_ascii_case("FACTDOUBLE") {
        Function::FactDouble
    } else if name.eq_ignore_ascii_case("GCD") {
        Function::Gcd
    } else if name.eq_ignore_ascii_case("GESTEP") {
        Function::GEstep
    } else if name.eq_ignore_ascii_case("LCM") {
        Function::Lcm
    } else if name.eq_ignore_ascii_case("MULTINOMIAL") {
        Function::Multinomial
    } else if name.eq_ignore_ascii_case("ODD") {
        Function::Odd
    } else {
        return None;
    })
}

pub(super) fn is_discrete_function(name: &str) -> bool {
    function(name).is_some()
}

/// Apply one discrete function after eager argument evaluation.
#[inline(never)]
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let Some(function) = function(name) else {
        return Err(EvaluationFailure::Unsupported(UnsupportedKind::Function));
    };
    match function {
        Function::Combin => apply_combin(evaluator, node),
        Function::Combina => apply_combina(evaluator, node),
        Function::Delta => apply_delta(evaluator, node),
        Function::Even | Function::Odd => apply_parity(evaluator, node, function),
        Function::Fact | Function::FactDouble => apply_factorial(evaluator, node, function),
        Function::Gcd | Function::Lcm | Function::Multinomial => {
            apply_sequence(evaluator, node, function)
        },
        Function::GEstep => apply_gestep(evaluator, node),
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
    evaluator.push_value(match result {
        Ok(value) if value.is_finite() => {
            WorkingValue::Number(if value == 0.0 { 0.0 } else { value })
        },
        Ok(_) => WorkingValue::Error(ScalarError::Number),
        Err(error) => WorkingValue::Error(error),
    })
}

fn apply_combin<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    if node.child_count() != 2 {
        return evaluator.finish_invalid_arity(node);
    }
    let right = evaluator.pop_value()?;
    let left = evaluator.pop_value()?;
    if let Some(error) = formula_error(&left) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(error) = formula_error(&right) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let n = match nonnegative_integer(super::to_number(left, evaluator)?) {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let r = match nonnegative_integer(super::to_number(right, evaluator)?) {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    charge_binomial(evaluator, n, r)?;
    push_uint(evaluator, binomial(n, r))
}

fn apply_combina<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
) -> EvaluationResult<()> {
    if node.child_count() != 2 {
        return evaluator.finish_invalid_arity(node);
    }
    let right = evaluator.pop_value()?;
    let left = evaluator.pop_value()?;
    if let Some(error) = formula_error(&left) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(error) = formula_error(&right) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let n = match nonnegative_integer(super::to_number(left, evaluator)?) {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let m = match nonnegative_integer(super::to_number(right, evaluator)?) {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    // Part 4 permits N=0,M=0 even though the displayed binomial expression
    // has a -1 in both slots.  The empty multiset has one combination.
    let result = if n.is_zero() && m.is_zero() {
        Ok(Uint1024::one())
    } else if n.cmp(&m).is_lt() {
        Err(ScalarError::Number)
    } else if m.is_zero() || n == Uint1024::one() {
        Ok(Uint1024::one())
    } else if m == Uint1024::one() {
        Ok(n)
    } else {
        let n_plus_m_minus_one = n
            .add(&m)
            .and_then(|value| value.sub_small(1))
            .ok_or(ScalarError::Number);
        match n_plus_m_minus_one {
            Ok(total) => {
                let Some(k) = n.sub_small(1) else {
                    return evaluator.push_value(WorkingValue::Error(ScalarError::Number));
                };
                charge_binomial(evaluator, total, k)?;
                binomial(total, k)
            },
            Err(error) => Err(error),
        }
    };
    push_uint(evaluator, result)
}

fn apply_delta<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    let count = node.child_count();
    if !(1..=2).contains(&count) {
        return evaluator.finish_invalid_arity(node);
    }
    let right = if count == 2 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let left = evaluator.pop_value()?;
    if let Some(error) = formula_error(&left) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(value) = &right {
        if let Some(error) = formula_error(value) {
            return evaluator.push_value(WorkingValue::Error(error));
        }
    }
    let left = match super::to_number(left, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let right = match right {
        Some(value) => match super::to_number(value, evaluator)? {
            Ok(value) => value,
            Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
        },
        None => 0.0,
    };
    evaluator.push_value(WorkingValue::Number(if left == right { 1.0 } else { 0.0 }))
}

fn apply_gestep<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    let count = node.child_count();
    if !(1..=2).contains(&count) {
        return evaluator.finish_invalid_arity(node);
    }
    let right = if count == 2 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let left = evaluator.pop_value()?;
    if let Some(error) = formula_error(&left) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(value) = &right {
        if let Some(error) = formula_error(value) {
            return evaluator.push_value(WorkingValue::Error(error));
        }
    }
    let left = match super::to_number(left, evaluator)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let right = match right {
        Some(value) => match super::to_number(value, evaluator)? {
            Ok(value) => value,
            Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
        },
        None => 0.0,
    };
    evaluator.push_value(WorkingValue::Number(if left >= right { 1.0 } else { 0.0 }))
}

fn apply_parity<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    function: Function,
) -> EvaluationResult<()> {
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
    let negative = value < 0.0;
    let magnitude = value.abs().ceil();
    let mut magnitude = match Uint1024::from_f64_integer(magnitude) {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let want_odd = function == Function::Odd;
    if magnitude.bit(0) != want_odd {
        magnitude = match magnitude.add_small(1) {
            Some(value) => value,
            None => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
        };
    }
    let result = match magnitude.to_f64_nearest_even() {
        Some(value) => {
            // A rounded binary64 value may lose the low bit of an exact
            // candidate above 2^53.  The profile admits that rounded value
            // when it preserves the requested parity; otherwise the exact
            // odd/even result is not representable and is #NUM!.
            let rounded_magnitude = match Uint1024::from_f64_integer(value) {
                Ok(value) => value,
                Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
            };
            if rounded_magnitude != magnitude && rounded_magnitude.bit(0) != want_odd {
                return evaluator.push_value(WorkingValue::Error(ScalarError::Number));
            }
            if negative { -value } else { value }
        },
        None => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
    };
    push_result(evaluator, Ok(result))
}

fn apply_factorial<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    function: Function,
) -> EvaluationResult<()> {
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
    let integer = match nonnegative_integer(Ok(value)) {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let result = if function == Function::Fact {
        factorial(integer, evaluator)
    } else {
        double_factorial(integer, evaluator)
    };
    push_uint(evaluator, result?)
}

fn apply_sequence<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    function: Function,
) -> EvaluationResult<()> {
    if node.child_count() == 0 {
        return evaluator.finish_invalid_arity(node);
    }
    reverse_value_tail(evaluator, node.child_count())?;
    let mut accumulator = DiscreteFold::new(match function {
        Function::Gcd => DiscreteFunction::Gcd,
        Function::Lcm => DiscreteFunction::Lcm,
        Function::Multinomial => DiscreteFunction::Multinomial,
        _ => {
            return Err(EvaluationFailure::InvalidExpression(
                "non-sequence discrete function reached reducer",
            ));
        },
    });
    for _ in 0..node.child_count() {
        let value = evaluator.pop_value()?;
        match value {
            WorkingValue::Error(error) => accumulator.push_formula_error(error),
            value => match super::to_number(value, evaluator)? {
                Ok(value) => {
                    // Admit the bounded kernel work before executing its
                    // fixed-width arithmetic.  The value bridge can use the
                    // same estimate through `DiscreteFold::work_for`.
                    let work = accumulator.work_for(value);
                    evaluator.charge_work(work as u64)?;
                    if let Err(error) = accumulator.push_number(value) {
                        accumulator.push_generated_error(error);
                    }
                },
                Err(error) => accumulator.push_generated_error(error),
            },
        }
    }
    push_result(evaluator, accumulator.finish())
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
                "discrete value stack underflow",
            ))?;
    evaluator.values[start..].reverse();
    Ok(())
}

fn nonnegative_integer(value: Result<f64, ScalarError>) -> Result<Uint1024, ScalarError> {
    let value = value?;
    if !value.is_finite() || value < 0.0 {
        return Err(ScalarError::Number);
    }
    Uint1024::from_f64_integer(value.floor())
}

fn push_uint(
    evaluator: &mut Evaluator<'_, '_, '_>,
    result: Result<Uint1024, ScalarError>,
) -> EvaluationResult<()> {
    push_result(
        evaluator,
        result.and_then(|value| value.to_f64_nearest_even().ok_or(ScalarError::Number)),
    )
}

fn factorial(
    value: Uint1024,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<Uint1024, ScalarError>> {
    let Some(value) = value.to_u64() else {
        return Ok(Err(ScalarError::Number));
    };
    // 171! is already beyond finite binary64.  This is a mathematical
    // boundary, not an evaluator loop cap; it avoids iterating over a huge
    // represented integer after the result is known to be unrepresentable.
    if value > 170 {
        return Ok(Err(ScalarError::Number));
    }
    let mut result = Uint1024::one();
    for factor in 2..=value {
        evaluator.charge_work(1)?;
        let Some(product) = result.mul(&Uint1024::from_u64(factor)) else {
            return Ok(Err(ScalarError::Number));
        };
        result = product;
    }
    Ok(Ok(result))
}

fn double_factorial(
    value: Uint1024,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<Uint1024, ScalarError>> {
    let Some(value) = value.to_u64() else {
        return Ok(Err(ScalarError::Number));
    };
    // 300!! is finite while 301!! is not.  As with FACT, this is derived
    // from the exact integer boundary and is not an arbitrary safety cap.
    if value > 300 {
        return Ok(Err(ScalarError::Number));
    }
    let mut result = Uint1024::one();
    let mut factor = value;
    while factor >= 2 {
        evaluator.charge_work(1)?;
        let Some(product) = result.mul(&Uint1024::from_u64(factor)) else {
            return Ok(Err(ScalarError::Number));
        };
        result = product;
        factor -= 2;
    }
    Ok(Ok(result))
}

fn charge_binomial(
    evaluator: &mut Evaluator<'_, '_, '_>,
    n: Uint1024,
    k: Uint1024,
) -> EvaluationResult<()> {
    if k.cmp(&n).is_gt() {
        return Ok(());
    }
    let complement = n.sub(&k);
    let rounds = if k.cmp(&complement).is_lt() {
        k
    } else {
        complement
    };
    if let Some(rounds) = rounds.to_u64()
        && rounds < 1024
    {
        evaluator.charge_work(rounds)?;
    }
    Ok(())
}

fn binomial(n: Uint1024, k: Uint1024) -> Result<Uint1024, ScalarError> {
    if k.cmp(&n).is_gt() {
        return Err(ScalarError::Number);
    }
    let complement = n.sub(&k);
    let r = if k.cmp(&complement).is_lt() {
        k
    } else {
        complement
    };
    if r.is_zero() {
        return Ok(Uint1024::one());
    }
    if r == Uint1024::one() {
        return Ok(n);
    }
    let Some(rounds) = r.to_u64() else {
        return Err(ScalarError::Number);
    };
    // If r is at least 1024, n >= 2r and C(n,r) >= 2^r, which cannot round
    // to a finite binary64 Number.  The recurrence therefore has a fixed
    // finite bound without imposing a host-specific argument cap.
    if rounds >= 1024 {
        return Err(ScalarError::Number);
    }
    let base = n.sub(&r);
    let mut result = Uint1024::one();
    for index in 1..=rounds {
        let numerator = base.add_small(index).ok_or(ScalarError::Number)?;
        let numerator_gcd = numerator.gcd_small(index);
        let numerator = numerator
            .div_small_exact(numerator_gcd)
            .ok_or(ScalarError::Number)?;
        let mut denominator = index / numerator_gcd;
        let result_gcd = result.gcd_small(denominator);
        result = result
            .div_small_exact(result_gcd)
            .ok_or(ScalarError::Number)?;
        denominator /= result_gcd;
        let product = result.mul(&numerator).ok_or(ScalarError::Number)?;
        // The recurrence is integral.  Retain this checked division rather
        // than relying on that invariant if the fixed-width state changes.
        result = if denominator == 1 {
            product
        } else {
            product
                .div_exact(&Uint1024::from_u64(denominator))
                .ok_or(ScalarError::Number)?
        };
    }
    Ok(result)
}

const FRACTION_BITS: usize = 1074;
const FRACTION_LIMBS: usize = FRACTION_BITS.div_ceil(BITS_PER_LIMB) + 1;

/// A binary fixed-point sum of fractions in the range `[0, 1)`, scaled by
/// `2^1074`.  Each admitted binary64 fraction is represented exactly.  The
/// extra limb holds the carry bit when a sequence's fractional parts cross
/// one; the carry is removed immediately and returned to the caller.
#[derive(Clone, Copy)]
struct FractionalSum {
    limbs: [u32; FRACTION_LIMBS],
}

impl FractionalSum {
    const fn zero() -> Self {
        Self {
            limbs: [0; FRACTION_LIMBS],
        }
    }

    fn bit(&self, index: usize) -> bool {
        index < FRACTION_LIMBS * BITS_PER_LIMB
            && (self.limbs[index / BITS_PER_LIMB] & (1_u32 << (index % BITS_PER_LIMB))) != 0
    }

    fn clear_bit(&mut self, index: usize) {
        if index < FRACTION_LIMBS * BITS_PER_LIMB {
            self.limbs[index / BITS_PER_LIMB] &= !(1_u32 << (index % BITS_PER_LIMB));
        }
    }

    fn add_fraction(&mut self, value: f64) -> bool {
        debug_assert!(value.is_finite() && value >= 0.0);
        let bits = value.to_bits();
        let exponent = ((bits >> 52) & 0x7ff) as i32;
        let mantissa = bits & ((1_u64 << 52) - 1);
        if exponent == 0x7ff {
            return false;
        }
        let (significand, shift) = if exponent == 0 {
            // Subnormal values are already an integer number of 2^-1074.
            (mantissa, 0)
        } else {
            let unbiased = exponent - 1023;
            if unbiased >= 52 {
                return false;
            }
            let significand = (1_u64 << 52) | mantissa;
            if unbiased < 0 {
                (significand, (unbiased + 1022) as usize)
            } else {
                let low_bits = (1_u64 << (52 - unbiased as u32)) - 1;
                (significand & low_bits, (unbiased + 1022) as usize)
            }
        };
        self.add_shifted(significand, shift);
        if self.bit(FRACTION_BITS) {
            self.clear_bit(FRACTION_BITS);
            true
        } else {
            false
        }
    }

    fn add_shifted(&mut self, value: u64, shift: usize) {
        // The significand has at most 53 bits and the scaled fraction has at
        // most bit 1073 set.  A bitwise carry loop is therefore fixed at 53
        // additions and at most the fixed 35-limb state.
        for bit in 0..=52 {
            if value & (1_u64 << bit) == 0 {
                continue;
            }
            let mut index = shift + bit;
            let mut carry = 1_u64;
            while carry != 0 && index < self.limbs.len() * BITS_PER_LIMB {
                let limb = index / BITS_PER_LIMB;
                let offset = index % BITS_PER_LIMB;
                let sum = u64::from(self.limbs[limb]) + (carry << offset);
                self.limbs[limb] = sum as u32;
                carry = sum >> 32;
                index = (limb + 1) * BITS_PER_LIMB;
            }
            debug_assert_eq!(carry, 0);
        }
    }
}

/// The sequence reducer kind shared with the value evaluator.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum DiscreteFunction {
    Gcd,
    Lcm,
    Multinomial,
}

/// Return the shared sequence kind for a discrete function name.
pub(super) fn sequence_function(name: &str) -> Option<DiscreteFunction> {
    Some(if name.eq_ignore_ascii_case("GCD") {
        DiscreteFunction::Gcd
    } else if name.eq_ignore_ascii_case("LCM") {
        DiscreteFunction::Lcm
    } else if name.eq_ignore_ascii_case("MULTINOMIAL") {
        DiscreteFunction::Multinomial
    } else {
        return None;
    })
}

/// Allocation-free streaming state for GCD, LCM, and MULTINOMIAL.
pub(super) struct DiscreteFold {
    function: DiscreteFunction,
    value: Uint1024,
    total: Uint1024,
    has_value: bool,
    overflow: bool,
    zero: bool,
    fraction: FractionalSum,
    fraction_carry: u64,
    formula_error: Option<ScalarError>,
    generated_error: Option<ScalarError>,
}

impl DiscreteFold {
    pub(super) fn new(function: DiscreteFunction) -> Self {
        let identity = if function == DiscreteFunction::Multinomial {
            Uint1024::one()
        } else {
            Uint1024::zero()
        };
        Self {
            function,
            value: identity,
            total: Uint1024::zero(),
            has_value: false,
            overflow: false,
            zero: false,
            fraction: FractionalSum::zero(),
            fraction_carry: 0,
            formula_error: None,
            generated_error: None,
        }
    }

    pub(super) fn push_formula_error(&mut self, error: ScalarError) {
        if self.formula_error.is_none() {
            self.formula_error = Some(error);
        }
    }

    pub(super) fn push_generated_error(&mut self, error: ScalarError) {
        if self.generated_error.is_none() {
            self.generated_error = Some(error);
        }
    }

    pub(super) fn push_number(&mut self, value: f64) -> Result<(), ScalarError> {
        if self.formula_error.is_some() {
            return Ok(());
        }
        let generated_before = self.generated_error;
        match self.function {
            DiscreteFunction::Gcd => match nonnegative_integer(Ok(value)) {
                Ok(value) => self.push_gcd(value),
                Err(error) => self.push_generated_error(error),
            },
            DiscreteFunction::Lcm => {
                // LCM's contract keeps the integer constraint separate from
                // INT conversion: a finite fraction is a domain error.
                if !value.is_finite() || value < 0.0 || value.fract() != 0.0 {
                    self.push_generated_error(ScalarError::Number);
                } else {
                    match Uint1024::from_f64_integer(value) {
                        Ok(value) => self.push_lcm(value),
                        Err(error) => self.push_generated_error(error),
                    }
                }
            },
            DiscreteFunction::Multinomial => self.push_multinomial_raw(value),
        }
        if generated_before.is_none() {
            if let Some(error) = self.generated_error {
                return Err(error);
            }
        }
        Ok(())
    }

    fn estimate_work(&self, value: f64) -> usize {
        let mut work = 1usize;
        match self.function {
            DiscreteFunction::Gcd | DiscreteFunction::Lcm => {
                let incoming_is_wide = Uint1024::from_f64_integer(value)
                    .ok()
                    .is_none_or(|value| value.to_u64().is_none());
                if incoming_is_wide || self.value.to_u64().is_none() {
                    // Wide operands use the bounded binary long-division
                    // path; ordinary u64 operands use the fast Euclidean
                    // path.  Charge the fixed-width path as one bounded
                    // reducer unit per possible bit step.
                    work = work.saturating_add(MAX_BITS * 2);
                }
            },
            DiscreteFunction::Multinomial => {
                // Decoding a binary64 fraction examines at most 53 bits.
                work = work.saturating_add(53);
                let mut fraction = self.fraction;
                if value.is_finite() && value >= 0.0 && fraction.add_fraction(value) {
                    // One future rising-product multiplication is added to
                    // the finish cost for each exact fractional carry.
                    work = work.saturating_add(1);
                }
                if let Ok(integer) = Uint1024::from_f64_integer(value.floor()) {
                    if let Some(total) = self.total.add(&integer) {
                        let complement = total.sub(&integer);
                        let rounds = if integer.cmp(&complement).is_lt() {
                            integer
                        } else {
                            complement
                        };
                        work =
                            work.saturating_add(rounds.to_u64().unwrap_or(1024).min(1024) as usize);
                    } else {
                        work = work.saturating_add(1024);
                    }
                }
            },
        }
        work
    }

    /// Return the bounded work admission required by one input.  Callers
    /// that own an evaluator should charge this value before `push_number`.
    pub(super) fn work_for(&self, value: f64) -> usize {
        self.estimate_work(value)
    }

    fn push_gcd(&mut self, value: Uint1024) {
        if !self.has_value {
            self.value = value;
            self.has_value = true;
        } else {
            self.value = self.value.gcd(&value);
        }
    }

    fn push_lcm(&mut self, value: Uint1024) {
        self.has_value = true;
        if value.is_zero() {
            // Zero is the absorbing LCM result.  Keep scanning later values
            // for source-order formula errors, but let this reset a prior
            // intermediate overflow as required by the mathematical result.
            self.zero = true;
            self.overflow = false;
            self.value = Uint1024::zero();
            return;
        }
        if self.zero || self.overflow {
            return;
        }
        if self.value.is_zero() {
            self.value = value;
            return;
        }
        let gcd = self.value.gcd(&value);
        let reduced = match self.value.div_exact(&gcd) {
            Some(value) => value,
            None => {
                self.push_generated_error(ScalarError::Number);
                return;
            },
        };
        match reduced.mul(&value) {
            Some(value) => self.value = value,
            None => self.overflow = true,
        }
    }

    fn push_multinomial_raw(&mut self, value: f64) {
        if !value.is_finite() || value < 0.0 {
            self.push_generated_error(ScalarError::Number);
            return;
        }
        if self.fraction.add_fraction(value) {
            self.fraction_carry = match self.fraction_carry.checked_add(1) {
                Some(value) => value,
                None => {
                    self.push_generated_error(ScalarError::Number);
                    return;
                },
            };
        }
        let value = match Uint1024::from_f64_integer(value.floor()) {
            Ok(value) => value,
            Err(error) => {
                self.push_generated_error(error);
                return;
            },
        };
        self.push_multinomial_integer(value);
    }

    fn push_multinomial_integer(&mut self, value: Uint1024) {
        if self.overflow || value.is_zero() {
            self.has_value = true;
            return;
        }
        let next_total = match self.total.add(&value) {
            Some(value) => value,
            None => {
                self.overflow = true;
                self.has_value = true;
                return;
            },
        };
        let factor = match binomial(next_total, value) {
            Ok(value) => value,
            Err(error) => {
                self.push_generated_error(error);
                self.has_value = true;
                return;
            },
        };
        self.total = next_total;
        self.has_value = true;
        match self.value.mul(&factor) {
            Some(value) => self.value = value,
            None => self.overflow = true,
        }
    }

    pub(super) fn finish(self) -> Result<f64, ScalarError> {
        if let Some(error) = self.formula_error.or(self.generated_error) {
            return Err(error);
        }
        if !self.has_value {
            return match self.function {
                DiscreteFunction::Multinomial => Ok(1.0),
                DiscreteFunction::Gcd | DiscreteFunction::Lcm => Err(ScalarError::Number),
            };
        }
        if self.function == DiscreteFunction::Lcm && self.zero {
            return Ok(0.0);
        }
        if self.overflow {
            return Err(ScalarError::Number);
        }
        let value = match self.function {
            DiscreteFunction::Gcd | DiscreteFunction::Lcm => self.value,
            DiscreteFunction::Multinomial => {
                // The raw sum is the integer parts' total plus every carry
                // from the exact fractional residue.  The missing numerator
                // factors are (T+1)..(T+c) = c! * C(T+c,c), so this rising
                // product completes the factorial ratio without materializing
                // the potentially huge raw sum.  If c >= 171, c! alone is
                // outside finite binary64, which proves #NUM before looping.
                if self.fraction_carry > 170 {
                    return Err(ScalarError::Number);
                }
                let mut value = self.value;
                for index in 1..=self.fraction_carry {
                    let factor = self.total.add_small(index).ok_or(ScalarError::Number)?;
                    value = value.mul(&factor).ok_or(ScalarError::Number)?;
                }
                value
            },
        };
        value.to_f64_nearest_even().ok_or(ScalarError::Number)
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
struct Uint1024 {
    limbs: [u32; MAX_LIMBS],
}

impl Uint1024 {
    const fn zero() -> Self {
        Self {
            limbs: [0; MAX_LIMBS],
        }
    }

    const fn one() -> Self {
        let mut value = Self::zero();
        value.limbs[0] = 1;
        value
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
        if highest >= MAX_BITS {
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
        if index < MAX_BITS {
            self.limbs[index / BITS_PER_LIMB] |= 1_u32 << (index % BITS_PER_LIMB);
        }
    }

    fn bit(self, index: usize) -> bool {
        index < MAX_BITS
            && (self.limbs[index / BITS_PER_LIMB] & (1_u32 << (index % BITS_PER_LIMB))) != 0
    }

    fn is_zero(self) -> bool {
        self.limbs.iter().all(|limb| *limb == 0)
    }

    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        for index in (0..MAX_LIMBS).rev() {
            match self.limbs[index].cmp(&other.limbs[index]) {
                std::cmp::Ordering::Equal => {},
                ordering => return ordering,
            }
        }
        std::cmp::Ordering::Equal
    }

    fn add(&self, other: &Self) -> Option<Self> {
        let mut result = Self::zero();
        let mut carry = 0_u64;
        for index in 0..MAX_LIMBS {
            let value = u64::from(self.limbs[index])
                .checked_add(u64::from(other.limbs[index]))?
                .checked_add(carry)?;
            result.limbs[index] = value as u32;
            carry = value >> 32;
        }
        (carry == 0).then_some(result)
    }

    fn add_small(&self, value: u64) -> Option<Self> {
        self.add(&Self::from_u64(value))
    }

    fn sub(&self, other: &Self) -> Self {
        debug_assert!(self.cmp(other).is_ge());
        let mut result = Self::zero();
        let mut borrow = 0_u64;
        for index in 0..MAX_LIMBS {
            let left = u64::from(self.limbs[index]);
            let right = u64::from(other.limbs[index]) + borrow;
            if left >= right {
                result.limbs[index] = (left - right) as u32;
                borrow = 0;
            } else {
                result.limbs[index] = ((1_u64 << 32) + left - right) as u32;
                borrow = 1;
            }
        }
        result
    }

    fn sub_small(&self, value: u64) -> Option<Self> {
        let other = Self::from_u64(value);
        (self.cmp(&other).is_ge()).then(|| self.sub(&other))
    }

    fn mul(&self, other: &Self) -> Option<Self> {
        if self.is_zero() || other.is_zero() {
            return Some(Self::zero());
        }
        let left_len = self.limb_len();
        let right_len = other.limb_len();
        let mut wide = [0_u32; MAX_LIMBS * 2];
        for left in 0..left_len {
            let mut carry = 0_u64;
            for right in 0..right_len {
                let index = left + right;
                let product = u64::from(self.limbs[left]) * u64::from(other.limbs[right]);
                let value = u128::from(wide[index]) + u128::from(product) + u128::from(carry);
                wide[index] = value as u32;
                carry = (value >> 32) as u64;
            }
            let mut index = left + right_len;
            while carry != 0 && index < wide.len() {
                let value = u64::from(wide[index]) + carry;
                wide[index] = value as u32;
                carry = value >> 32;
                index += 1;
            }
            if carry != 0 {
                return None;
            }
        }
        if wide[MAX_LIMBS..].iter().any(|limb| *limb != 0) {
            return None;
        }
        let mut result = Self::zero();
        result.limbs.copy_from_slice(&wide[..MAX_LIMBS]);
        Some(result)
    }

    fn limb_len(&self) -> usize {
        self.limbs
            .iter()
            .rposition(|limb| *limb != 0)
            .map_or(0, |index| index + 1)
    }

    fn div_rem(&self, divisor: &Self) -> Option<(Self, Self)> {
        if divisor.is_zero() {
            return None;
        }
        let mut quotient = Self::zero();
        let mut remainder = Self::zero();
        let Some(highest) = self.highest_bit() else {
            return Some((quotient, remainder));
        };
        for index in (0..=highest).rev() {
            let carry = remainder.shl1_add(self.bit(index));
            if carry || remainder.cmp(divisor).is_ge() {
                remainder = remainder.sub_wrapping(divisor);
                quotient.set_bit(index);
            }
        }
        Some((quotient, remainder))
    }

    fn div_exact(&self, divisor: &Self) -> Option<Self> {
        let (quotient, remainder) = self.div_rem(divisor)?;
        remainder.is_zero().then_some(quotient)
    }

    fn shl1_add(&mut self, bit: bool) -> bool {
        let mut carry = u64::from(bit);
        for limb in &mut self.limbs {
            let value = (u64::from(*limb) << 1) | carry;
            *limb = value as u32;
            carry = value >> 32;
        }
        carry != 0
    }

    fn sub_wrapping(&self, other: &Self) -> Self {
        let mut result = Self::zero();
        let mut borrow = 0_u64;
        for index in 0..MAX_LIMBS {
            let left = u64::from(self.limbs[index]);
            let right = u64::from(other.limbs[index]) + borrow;
            if left >= right {
                result.limbs[index] = (left - right) as u32;
                borrow = 0;
            } else {
                result.limbs[index] = ((1_u64 << 32) + left - right) as u32;
                borrow = 1;
            }
        }
        result
    }

    fn gcd(&self, other: &Self) -> Self {
        if let (Some(left), Some(right)) = (self.to_u64(), other.to_u64()) {
            return Self::from_u64(gcd_u64(left, right));
        }
        let mut left = *self;
        let mut right = *other;
        for _ in 0..=(MAX_BITS * 2) {
            if right.is_zero() {
                return left;
            }
            let Some((_, remainder)) = left.div_rem(&right) else {
                return Self::zero();
            };
            left = right;
            right = remainder;
        }
        Self::zero()
    }

    fn gcd_small(&self, divisor: u64) -> u64 {
        if divisor == 0 {
            return 0;
        }
        let remainder = self.rem_small(divisor);
        gcd_u64(remainder, divisor)
    }

    fn rem_small(&self, divisor: u64) -> u64 {
        let mut remainder = 0_u64;
        for limb in self.limbs.iter().rev() {
            remainder = ((remainder << 32) + u64::from(*limb)) % divisor;
        }
        remainder
    }

    fn div_small_exact(&self, divisor: u64) -> Option<Self> {
        if divisor == 0 {
            return None;
        }
        let mut result = Self::zero();
        let mut remainder = 0_u64;
        for (index, limb) in self.limbs.iter().enumerate().rev() {
            let value = (remainder << 32) + u64::from(*limb);
            result.limbs[index] = (value / divisor) as u32;
            remainder = value % divisor;
        }
        (remainder == 0).then_some(result)
    }

    fn to_u64(self) -> Option<u64> {
        if self.limbs[2..].iter().any(|limb| *limb != 0) {
            return None;
        }
        Some((u64::from(self.limbs[1]) << 32) | u64::from(self.limbs[0]))
    }

    fn highest_bit(self) -> Option<usize> {
        for (index, limb) in self.limbs.iter().enumerate().rev() {
            if *limb != 0 {
                return Some(index * BITS_PER_LIMB + (31 - limb.leading_zeros() as usize));
            }
        }
        None
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

    fn to_f64_nearest_even(self) -> Option<f64> {
        let Some(highest) = self.highest_bit() else {
            return Some(0.0);
        };
        if highest <= 52 {
            let mut result = 0.0;
            for bit in (0..=highest).rev() {
                result = result * 2.0 + f64::from(self.bit(bit));
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

fn gcd_u64(mut left: u64, mut right: u64) -> u64 {
    while right != 0 {
        let remainder = left % right;
        left = right;
        right = remainder;
    }
    left
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn exact_integer_state_preserves_large_additions() {
        let total = Uint1024::from_u64(1_u64 << 53)
            .add_small(1)
            .and_then(|value| value.add_small(1))
            .unwrap();
        assert_eq!(total.to_u64(), Some((1_u64 << 53) + 2));
    }

    #[test]
    fn exact_reducers_handle_zero_and_large_factors() {
        let mut lcm = DiscreteFold::new(DiscreteFunction::Lcm);
        lcm.push_number(48.0).unwrap();
        lcm.push_number(18.0).unwrap();
        lcm.push_number(0.0).unwrap();
        assert_eq!(lcm.finish().unwrap(), 0.0);

        let mut multinomial = DiscreteFold::new(DiscreteFunction::Multinomial);
        multinomial.push_number(2_f64.powi(53)).unwrap();
        multinomial.push_number(1.0).unwrap();
        multinomial.push_number(1.0).unwrap();
        let expected = ((1_u128 << 53) + 1) * ((1_u128 << 53) + 2);
        assert_eq!(multinomial.finish().unwrap(), expected as f64);
    }

    #[test]
    fn parity_uses_exact_represented_integer_bits() {
        let value = Uint1024::from_f64_integer(2_f64.powi(53)).unwrap();
        assert!(!value.bit(0));
        assert_eq!(
            value.add_small(1).unwrap().to_f64_nearest_even(),
            Some(2_f64.powi(53))
        );
    }
}
