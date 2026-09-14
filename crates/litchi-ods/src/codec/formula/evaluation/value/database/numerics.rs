//! Allocation-free numeric kernels for the ODF database-function family.
//!
//! The database evaluator owns selection, conversion, error ordering, and
//! resource accounting.  This module only receives finite `f64` numbers and
//! keeps the state needed by the numeric aggregates.  In particular, the sum
//! state is a fixed binary accumulator rather than an `f64`: all finite binary
//! floating-point values are integral multiples of `2^-1074`, so a pair of
//! 34-limb magnitudes can represent a bounded sum without losing cancellation
//! between a subnormal and values near `f64::MAX`.  The evaluator still has to
//! charge each call to [`NumericAggregate::push_number`].

use super::super::super::ScalarError;
use std::cmp::Ordering;

// A finite f64 has at most 2,098 significant bits when expressed as an
// integer multiple of 2^-1074 (the high bit of MAX is bit 2,097).  The
// database evaluator has a finite input-cell limit, so a 34-limb magnitude
// also leaves room for a carry from the selected-cell count without using a
// heap allocation.  Two magnitudes occupy 544 bytes.
const SUM_LIMBS: usize = 34;
const SUM_BITS: usize = SUM_LIMBS * 64;
const SIGNIFICAND_MASK: u64 = (1_u64 << 52) - 1;
const SIGNIFICAND_HIDDEN: u64 = 1_u64 << 52;
const SIGN_BIT: u64 = 1_u64 << 63;

/// The numeric operation selected by one database-function invocation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum NumericOperation {
    /// Count selected Number cells (`DCOUNT` with a field).
    Count,
    /// Add selected Number cells (`DSUM`).
    Sum,
    /// Average selected Number cells (`DAVERAGE`).
    Average,
    /// Multiply selected Number cells (`DPRODUCT`).
    Product,
    /// Select the smallest Number (`DMIN`).
    Minimum,
    /// Select the largest Number (`DMAX`).
    Maximum,
    /// Sample variance (`DVAR`).
    SampleVariance,
    /// Population variance (`DVARP`).
    PopulationVariance,
    /// Sample standard deviation (`DSTDEV`).
    SampleStandardDeviation,
    /// Population standard deviation (`DSTDEVP`).
    PopulationStandardDeviation,
}

#[derive(Clone, Copy, Debug)]
enum NumericState {
    Count(u64),
    Sum(ExactSum),
    Average(ExactSum),
    Product(ProductAccumulator),
    Extrema(ExtremaAccumulator, bool),
    Variance(VarianceAccumulator, VarianceKind),
}

#[derive(Clone, Copy, Debug)]
enum VarianceKind {
    SampleVariance,
    PopulationVariance,
    SampleStandardDeviation,
    PopulationStandardDeviation,
}

/// The selected finite numeric values and the one aggregate state needed by
/// the current database function.
///
/// Construct this with [`Self::new`] for the function being evaluated.  This
/// is deliberately operation-specific: a `DSUM` scan never updates product,
/// extrema, or variance state, and therefore cannot inherit an unrelated
/// kernel's overflow refusal.  Empty selections are represented by `None` in
/// the result methods; the caller applies each function's explicit
/// empty-selection profile.
#[derive(Clone, Copy, Debug)]
pub(crate) struct NumericAggregate {
    state: NumericState,
}

impl Default for NumericAggregate {
    fn default() -> Self {
        Self::new(NumericOperation::Sum)
    }
}

impl NumericAggregate {
    /// Create a state for one database numeric operation.
    pub(crate) fn new(operation: NumericOperation) -> Self {
        let state = match operation {
            NumericOperation::Count => NumericState::Count(0),
            NumericOperation::Sum => NumericState::Sum(ExactSum::default()),
            NumericOperation::Average => NumericState::Average(ExactSum::default()),
            NumericOperation::Product => NumericState::Product(ProductAccumulator::default()),
            NumericOperation::Minimum => {
                NumericState::Extrema(ExtremaAccumulator::default(), false)
            },
            NumericOperation::Maximum => NumericState::Extrema(ExtremaAccumulator::default(), true),
            NumericOperation::SampleVariance => {
                NumericState::Variance(VarianceAccumulator::default(), VarianceKind::SampleVariance)
            },
            NumericOperation::PopulationVariance => NumericState::Variance(
                VarianceAccumulator::default(),
                VarianceKind::PopulationVariance,
            ),
            NumericOperation::SampleStandardDeviation => NumericState::Variance(
                VarianceAccumulator::default(),
                VarianceKind::SampleStandardDeviation,
            ),
            NumericOperation::PopulationStandardDeviation => NumericState::Variance(
                VarianceAccumulator::default(),
                VarianceKind::PopulationStandardDeviation,
            ),
        };
        Self { state }
    }

    /// Add one selected Number to the active numeric aggregate.
    ///
    /// The caller is responsible for filtering range Text, Logical, Empty,
    /// and Error cells according to the function's `NumberSequence` rules.
    /// Non-finite values are rejected as a formula-level `#NUM!` value.  The
    /// operation performs no allocation.
    pub(crate) fn push_number(&mut self, value: f64) -> Result<(), ScalarError> {
        if !value.is_finite() {
            return Err(ScalarError::Number);
        }
        match &mut self.state {
            NumericState::Count(count) => {
                *count = count.checked_add(1).ok_or(ScalarError::Number)?;
                Ok(())
            },
            NumericState::Sum(sum) | NumericState::Average(sum) => sum.push(value),
            NumericState::Product(product) => product.push(value),
            NumericState::Extrema(extrema, _) => extrema.push(value),
            NumericState::Variance(variance, _) => variance.push(value),
        }
    }

    /// Number of values accepted by [`Self::push_number`].
    #[must_use]
    pub(crate) const fn count(&self) -> u64 {
        match self.state {
            NumericState::Count(count) => count,
            NumericState::Sum(sum) | NumericState::Average(sum) => sum.count(),
            NumericState::Product(product) => product.count,
            NumericState::Extrema(extrema, _) => extrema.count,
            NumericState::Variance(variance, _) => variance.count,
        }
    }

    /// Return the selected operation's result.
    pub(crate) fn result(&self) -> Result<Option<f64>, ScalarError> {
        match self.state {
            NumericState::Count(count) => Ok(Some(count as f64)),
            NumericState::Sum(sum) => sum.result(),
            NumericState::Average(sum) => sum.average(),
            NumericState::Product(product) => product.result(),
            NumericState::Extrema(extrema, maximum) => Ok(extrema.result(maximum)),
            NumericState::Variance(variance, kind) => match kind {
                VarianceKind::SampleVariance => variance.variance(true),
                VarianceKind::PopulationVariance => variance.variance(false),
                VarianceKind::SampleStandardDeviation => variance.standard_deviation(true),
                VarianceKind::PopulationStandardDeviation => variance.standard_deviation(false),
            },
        }
    }
}

/// Exact signed sum in units of the smallest positive f64.
#[derive(Clone, Copy, Debug)]
struct ExactSum {
    positive: [u64; SUM_LIMBS],
    negative: [u64; SUM_LIMBS],
    count: u64,
}

impl Default for ExactSum {
    fn default() -> Self {
        Self {
            positive: [0; SUM_LIMBS],
            negative: [0; SUM_LIMBS],
            count: 0,
        }
    }
}

impl ExactSum {
    fn push(&mut self, value: f64) -> Result<(), ScalarError> {
        if !value.is_finite() {
            return Err(ScalarError::Number);
        }
        let next_count = self.count.checked_add(1).ok_or(ScalarError::Number)?;
        let bits = value.to_bits();
        let raw_exponent = ((bits >> 52) & 0x7ff) as usize;
        let fraction = bits & SIGNIFICAND_MASK;
        let (significand, shift) = if raw_exponent == 0 {
            // Zero is represented by an empty magnitude.  It still counts as
            // a selected Number for AVERAGE and the variance constraints.
            (fraction, 0_usize)
        } else {
            // For a normal value M * 2^(E-52), the exponent of the integer
            // multiple of 2^-1074 is E+1022 = raw_exponent-1.
            (SIGNIFICAND_HIDDEN | fraction, raw_exponent - 1)
        };
        if significand != 0 {
            let target = if bits & SIGN_BIT == 0 {
                &mut self.positive
            } else {
                &mut self.negative
            };
            add_shifted(target, significand, shift)?;
        }
        self.count = next_count;
        Ok(())
    }

    const fn count(&self) -> u64 {
        self.count
    }

    fn result(&self) -> Result<Option<f64>, ScalarError> {
        if self.count == 0 {
            return Ok(None);
        }
        let (magnitude, negative) = self.signed_magnitude();
        Ok(Some(integer_to_f64(&magnitude, negative, false)?))
    }

    fn average(&self) -> Result<Option<f64>, ScalarError> {
        if self.count == 0 {
            return Ok(None);
        }
        let (magnitude, negative) = self.signed_magnitude();
        Ok(Some(integer_ratio_to_f64(
            &magnitude, self.count, negative,
        )?))
    }

    fn signed_magnitude(&self) -> ([u64; SUM_LIMBS], bool) {
        match compare_limbs(&self.positive, &self.negative) {
            Ordering::Greater => (subtract_limbs(&self.positive, &self.negative), false),
            Ordering::Less => (subtract_limbs(&self.negative, &self.positive), true),
            Ordering::Equal => ([0; SUM_LIMBS], false),
        }
    }
}

/// Add a 53-bit significand at a binary bit offset.
fn add_shifted(
    limbs: &mut [u64; SUM_LIMBS],
    significand: u64,
    shift: usize,
) -> Result<(), ScalarError> {
    let limb_shift = shift / 64;
    let bit_shift = shift % 64;
    if limb_shift >= SUM_LIMBS {
        return Err(ScalarError::Number);
    }

    let wide = (significand as u128) << bit_shift;
    add_limb(limbs, limb_shift, wide as u64)?;
    let high = (wide >> 64) as u64;
    if high != 0 {
        add_limb(limbs, limb_shift + 1, high)?;
    }
    Ok(())
}

fn add_limb(
    limbs: &mut [u64; SUM_LIMBS],
    mut index: usize,
    mut value: u64,
) -> Result<(), ScalarError> {
    while value != 0 {
        let limb = limbs.get_mut(index).ok_or(ScalarError::Number)?;
        let (sum, overflow) = limb.overflowing_add(value);
        *limb = sum;
        value = u64::from(overflow);
        index = index.checked_add(1).ok_or(ScalarError::Number)?;
    }
    Ok(())
}

fn compare_limbs(left: &[u64; SUM_LIMBS], right: &[u64; SUM_LIMBS]) -> Ordering {
    for index in (0..SUM_LIMBS).rev() {
        match left[index].cmp(&right[index]) {
            Ordering::Equal => {},
            ordering => return ordering,
        }
    }
    Ordering::Equal
}

fn subtract_limbs(larger: &[u64; SUM_LIMBS], smaller: &[u64; SUM_LIMBS]) -> [u64; SUM_LIMBS] {
    let mut result = [0_u64; SUM_LIMBS];
    let mut borrow = false;
    for index in 0..SUM_LIMBS {
        let (difference, first_borrow) = larger[index].overflowing_sub(smaller[index]);
        let (difference, second_borrow) = difference.overflowing_sub(u64::from(borrow));
        result[index] = difference;
        borrow = first_borrow || second_borrow;
    }
    result
}

fn highest_bit(limbs: &[u64; SUM_LIMBS]) -> Option<usize> {
    for index in (0..SUM_LIMBS).rev() {
        if limbs[index] != 0 {
            return Some(index * 64 + (63 - limbs[index].leading_zeros() as usize));
        }
    }
    None
}

fn bit(limbs: &[u64; SUM_LIMBS], index: usize) -> bool {
    index < SUM_BITS && (limbs[index / 64] & (1_u64 << (index % 64))) != 0
}

fn any_below(limbs: &[u64; SUM_LIMBS], bits: usize) -> bool {
    let full_limbs = (bits / 64).min(SUM_LIMBS);
    if limbs[..full_limbs].iter().any(|limb| *limb != 0) {
        return true;
    }
    if full_limbs >= SUM_LIMBS {
        return false;
    }
    let remainder = bits % 64;
    remainder != 0 && limbs[full_limbs] & ((1_u64 << remainder) - 1) != 0
}

fn top_significand(limbs: &[u64; SUM_LIMBS], high: usize) -> u64 {
    if high < 52 {
        return limbs[0];
    }
    let mut result = 0_u64;
    for offset in 0..=52 {
        if bit(limbs, high - offset) {
            result |= 1_u64 << (52 - offset);
        }
    }
    result
}

/// Convert an integer multiple of `2^-1074` to a rounded finite f64.
fn integer_to_f64(
    magnitude: &[u64; SUM_LIMBS],
    negative: bool,
    extra_sticky: bool,
) -> Result<f64, ScalarError> {
    let Some(high) = highest_bit(magnitude) else {
        return Ok(if negative { -0.0 } else { 0.0 });
    };

    let mut significand = top_significand(magnitude, high);
    let right_shift = high.saturating_sub(52);
    if right_shift != 0 {
        let guard = bit(magnitude, right_shift - 1);
        let lower = any_below(magnitude, right_shift - 1) || extra_sticky;
        if guard && (lower || significand & 1 != 0) {
            significand = significand.checked_add(1).ok_or(ScalarError::Number)?;
        }
    }

    let mut exponent = high as i64 - 1074;
    if significand == 1_u64 << 53 {
        significand >>= 1;
        exponent = exponent.checked_add(1).ok_or(ScalarError::Number)?;
    }
    if exponent > 1023 {
        return Err(ScalarError::Number);
    }

    let sign = if negative { SIGN_BIT } else { 0 };
    if exponent >= -1022 {
        let exponent_bits = u64::try_from(exponent + 1023).map_err(|_| ScalarError::Number)?;
        let fraction = significand
            .checked_sub(SIGNIFICAND_HIDDEN)
            .ok_or(ScalarError::Number)?;
        return Ok(f64::from_bits(sign | (exponent_bits << 52) | fraction));
    }

    // For an exact integer sum, values below 2^-1022 are already integral
    // subnormal quanta; `extra_sticky` is only used for a divided average and
    // is handled by `integer_ratio_to_f64` before reaching this branch.
    Ok(f64::from_bits(sign | significand))
}

fn divide_limbs(numerator: &[u64; SUM_LIMBS], denominator: u64) -> ([u64; SUM_LIMBS], u64) {
    debug_assert!(denominator != 0);
    let mut quotient = [0_u64; SUM_LIMBS];
    let mut remainder = 0_u64;
    for index in (0..SUM_LIMBS).rev() {
        let wide = ((remainder as u128) << 64) | numerator[index] as u128;
        quotient[index] = (wide / denominator as u128) as u64;
        remainder = (wide % denominator as u128) as u64;
    }
    (quotient, remainder)
}

fn rounded_up(remainder: u64, denominator: u64, low_is_odd: bool) -> bool {
    if remainder == 0 {
        return false;
    }
    let complement = denominator - remainder;
    remainder > complement || (remainder == complement && low_is_odd)
}

fn increment_limbs(limbs: &mut [u64; SUM_LIMBS]) -> Result<(), ScalarError> {
    let mut carry = true;
    for limb in limbs.iter_mut() {
        if !carry {
            break;
        }
        let (value, overflow) = limb.overflowing_add(1);
        *limb = value;
        carry = overflow;
    }
    if carry {
        Err(ScalarError::Number)
    } else {
        Ok(())
    }
}

/// Convert a signed fixed-point magnitude divided by a positive cell count.
///
/// Division is done in the same fixed-size limb representation.  For normal
/// results, the nonzero integer remainder is sufficient as a sticky bit for
/// the final nearest-even conversion.  At the subnormal boundary the exact
/// remainder is compared with one half of a quantum so an average of tiny
/// values does not disappear merely because the unreduced sum is too large
/// for a direct `f64` conversion.
fn integer_ratio_to_f64(
    magnitude: &[u64; SUM_LIMBS],
    denominator: u64,
    negative: bool,
) -> Result<f64, ScalarError> {
    if denominator == 0 {
        return Err(ScalarError::Number);
    }
    if highest_bit(magnitude).is_none() {
        return Ok(if negative { -0.0 } else { 0.0 });
    }
    let (mut quotient, remainder) = divide_limbs(magnitude, denominator);
    if highest_bit(&quotient).is_none() {
        let rounded = rounded_up(remainder, denominator, false);
        return Ok(if rounded {
            f64::from_bits((if negative { SIGN_BIT } else { 0 }) | 1)
        } else if negative {
            -0.0
        } else {
            0.0
        });
    }

    let quotient_high = highest_bit(&quotient).ok_or(ScalarError::Number)?;
    if quotient_high <= 52 && rounded_up(remainder, denominator, quotient[0] & 1 != 0) {
        increment_limbs(&mut quotient)?;
        return integer_to_f64(&quotient, negative, false);
    }
    integer_to_f64(&quotient, negative, remainder != 0)
}

/// Mantissa/exponent product state.  Mantissas remain in `[1, 2)` and the
/// exponent is kept wide until the final conversion, so an intermediate
/// product need not overflow simply because the final result is zero or
/// subnormal.
#[derive(Clone, Copy, Debug)]
struct ProductAccumulator {
    mantissa: f64,
    exponent: i64,
    negative: bool,
    count: u64,
    zero: bool,
}

impl Default for ProductAccumulator {
    fn default() -> Self {
        Self {
            mantissa: 1.0,
            exponent: 0,
            negative: false,
            count: 0,
            zero: false,
        }
    }
}

impl ProductAccumulator {
    fn push(&mut self, value: f64) -> Result<(), ScalarError> {
        if !value.is_finite() {
            return Err(ScalarError::Number);
        }
        let next_count = self.count.checked_add(1).ok_or(ScalarError::Number)?;
        self.negative ^= value.is_sign_negative();
        if value == 0.0 {
            self.zero = true;
            self.count = next_count;
            return Ok(());
        }

        let (value_mantissa, value_exponent) = binary_parts(value.abs());
        let mut mantissa = self.mantissa * value_mantissa;
        let mut exponent = self
            .exponent
            .checked_add(i64::from(value_exponent))
            .ok_or(ScalarError::Number)?;
        if mantissa >= 2.0 {
            mantissa *= 0.5;
            exponent = exponent.checked_add(1).ok_or(ScalarError::Number)?;
        }
        if !mantissa.is_finite() {
            return Err(ScalarError::Number);
        }
        self.mantissa = mantissa;
        self.exponent = exponent;
        self.count = next_count;
        Ok(())
    }

    fn result(&self) -> Result<Option<f64>, ScalarError> {
        if self.count == 0 {
            return Ok(None);
        }
        if self.zero {
            return Ok(Some(if self.negative { -0.0 } else { 0.0 }));
        }
        let result = scale_binary(
            self.mantissa
                .copysign(if self.negative { -1.0 } else { 1.0 }),
            self.exponent,
        );
        if result.is_finite() {
            Ok(Some(result))
        } else {
            Err(ScalarError::Number)
        }
    }
}

fn binary_parts(value: f64) -> (f64, i32) {
    debug_assert!(value.is_finite() && value > 0.0);
    let bits = value.to_bits();
    let raw_exponent = ((bits >> 52) & 0x7ff) as i32;
    if raw_exponent == 0 {
        // 2^54 moves every nonzero subnormal into the normal range while the
        // recursive call keeps a single exact normalization path.
        let (mantissa, exponent) = binary_parts(value * 18_014_398_509_481_984.0);
        return (mantissa, exponent - 54);
    }
    let fraction = bits & SIGNIFICAND_MASK;
    let mantissa = f64::from_bits((1023_u64 << 52) | fraction);
    (mantissa, raw_exponent - 1023)
}

fn scale_binary(value: f64, exponent: i64) -> f64 {
    debug_assert!(value.is_finite() && value.abs() >= 1.0 && value.abs() < 2.0);
    let bits = value.to_bits();
    let sign = bits & SIGN_BIT;
    let fraction = bits & SIGNIFICAND_MASK;
    if exponent > 1023 {
        return f64::from_bits(sign | (0x7ff_u64 << 52));
    }
    if exponent >= -1022 {
        let target_exponent = (exponent + 1023) as u64;
        return f64::from_bits(sign | (target_exponent << 52) | fraction);
    }

    let right_shift = (-1022_i64).saturating_sub(exponent);
    if right_shift >= 64 {
        return f64::from_bits(sign);
    }
    let right_shift = right_shift as u32;
    let mask = (1_u64 << right_shift) - 1;
    let significand = SIGNIFICAND_HIDDEN | fraction;
    let mut subnormal = significand >> right_shift;
    let remainder = significand & mask;
    let halfway = 1_u64 << (right_shift - 1);
    if remainder > halfway || (remainder == halfway && subnormal & 1 != 0) {
        subnormal += 1;
    }
    if subnormal >= SIGNIFICAND_HIDDEN {
        return f64::from_bits(sign | SIGNIFICAND_HIDDEN);
    }
    f64::from_bits(sign | subnormal)
}

/// Stable first-value-offset moment state.
///
/// Values are normalized by the largest absolute value seen so far.  The
/// state retains the first value as an offset and accumulates compensated
/// `sum(x - first)` and `sum((x - first)^2)`.  This avoids both
/// `x - origin` overflow for opposite-sign values near `f64::MAX` and the
/// factor-of-two error that a two-point Welford update can get when the mean
/// rounds back to the first value.  When a larger scale arrives, the offset
/// and both moments are rescaled before the next difference is formed.  The
/// final square of the scale is formed in mantissa/exponent form.
#[derive(Clone, Copy, Debug, Default)]
struct VarianceAccumulator {
    count: u64,
    first: f64,
    scale: f64,
    origin: f64,
    sum: f64,
    sum_compensation: f64,
    squares: f64,
    squares_compensation: f64,
}

impl VarianceAccumulator {
    fn push(&mut self, value: f64) -> Result<(), ScalarError> {
        if !value.is_finite() {
            return Err(ScalarError::Number);
        }
        let next_count = self.count.checked_add(1).ok_or(ScalarError::Number)?;
        let absolute = value.abs();
        if self.count == 0 {
            self.count = 1;
            self.first = value;
            self.scale = absolute;
            self.origin = if absolute == 0.0 {
                0.0
            } else {
                value / absolute
            };
            self.sum = 0.0;
            self.sum_compensation = 0.0;
            self.squares = 0.0;
            self.squares_compensation = 0.0;
            return Ok(());
        }

        let mut scale = self.scale;
        let mut origin = self.origin;
        let mut sum = self.sum;
        let mut sum_compensation = self.sum_compensation;
        let mut squares = self.squares;
        let mut squares_compensation = self.squares_compensation;
        if absolute > scale {
            let ratio = if scale == 0.0 { 0.0 } else { scale / absolute };
            let ratio_squared = ratio * ratio;
            origin *= ratio;
            sum *= ratio;
            sum_compensation *= ratio;
            squares *= ratio_squared;
            squares_compensation *= ratio_squared;
            scale = absolute;
        }
        let normalized = if scale == 0.0 { 0.0 } else { value / scale };
        // Sterbenz-exact subtraction preserves an adjacent representable
        // delta (for example, `nextafter(1e150, +inf) - 1e150`) before the
        // result is scaled.  Opposite-sign values near MAX can overflow this
        // subtraction; use the already normalized origin in that case.
        let difference = if scale == 0.0 {
            0.0
        } else {
            match value - self.first {
                raw if raw.is_finite() => raw / scale,
                _ => normalized - origin,
            }
        };
        let square = difference * difference;
        compensated_add(&mut sum, &mut sum_compensation, difference);
        compensated_add(&mut squares, &mut squares_compensation, square);
        if !origin.is_finite()
            || !sum.is_finite()
            || !sum_compensation.is_finite()
            || !squares.is_finite()
            || !squares_compensation.is_finite()
        {
            return Err(ScalarError::Number);
        }
        self.count = next_count;
        self.scale = scale;
        self.origin = origin;
        self.sum = sum;
        self.sum_compensation = sum_compensation;
        self.squares = squares;
        self.squares_compensation = squares_compensation;
        Ok(())
    }

    fn variance(&self, sample: bool) -> Result<Option<f64>, ScalarError> {
        let denominator = self.denominator(sample)?;
        let Some(denominator) = denominator else {
            return Ok(None);
        };
        if self.scale == 0.0 {
            return Ok(Some(0.0));
        }
        let sum = self.sum + self.sum_compensation;
        let squares = self.squares + self.squares_compensation;
        let normalized_second_moment = sum * sum / self.count as f64;
        if !sum.is_finite() || !squares.is_finite() || !normalized_second_moment.is_finite() {
            return Err(ScalarError::Number);
        }
        let centered = (squares - normalized_second_moment).max(0.0);
        if !centered.is_finite() {
            return Err(ScalarError::Number);
        }
        let normalized = centered / denominator;
        if !normalized.is_finite() || normalized < 0.0 {
            return Err(ScalarError::Number);
        }
        Ok(Some(scale_product(self.scale, self.scale, normalized)?))
    }

    fn standard_deviation(&self, sample: bool) -> Result<Option<f64>, ScalarError> {
        let denominator = self.denominator(sample)?;
        let Some(denominator) = denominator else {
            return Ok(None);
        };
        if self.scale == 0.0 {
            return Ok(Some(0.0));
        }
        let sum = self.sum + self.sum_compensation;
        let squares = self.squares + self.squares_compensation;
        let normalized_second_moment = sum * sum / self.count as f64;
        if !sum.is_finite() || !squares.is_finite() || !normalized_second_moment.is_finite() {
            return Err(ScalarError::Number);
        }
        let centered = (squares - normalized_second_moment).max(0.0);
        if !centered.is_finite() {
            return Err(ScalarError::Number);
        }
        let normalized = centered / denominator;
        if !normalized.is_finite() || normalized < 0.0 {
            return Err(ScalarError::Number);
        }
        let result = self.scale * normalized.sqrt();
        if result.is_finite() {
            Ok(Some(result))
        } else {
            Err(ScalarError::Number)
        }
    }

    fn denominator(&self, sample: bool) -> Result<Option<f64>, ScalarError> {
        if self.count == 0 {
            return Ok(None);
        }
        if sample {
            if self.count < 2 {
                return Err(ScalarError::Value);
            }
            Ok(Some((self.count - 1) as f64))
        } else {
            Ok(Some(self.count as f64))
        }
    }
}

fn compensated_add(sum: &mut f64, compensation: &mut f64, value: f64) {
    let next = *sum + value;
    let correction = if sum.abs() >= value.abs() {
        (*sum - next) + value
    } else {
        (value - next) + *sum
    };
    *sum = next;
    *compensation += correction;
}

#[derive(Clone, Copy, Debug, Default)]
struct ExtremaAccumulator {
    count: u64,
    minimum: Option<f64>,
    maximum: Option<f64>,
}

impl ExtremaAccumulator {
    fn push(&mut self, value: f64) -> Result<(), ScalarError> {
        self.count = self.count.checked_add(1).ok_or(ScalarError::Number)?;
        self.minimum = Some(match self.minimum {
            Some(current) => current.min(value),
            None => value,
        });
        self.maximum = Some(match self.maximum {
            Some(current) => current.max(value),
            None => value,
        });
        Ok(())
    }

    const fn result(&self, maximum: bool) -> Option<f64> {
        if self.count == 0 {
            None
        } else if maximum {
            self.maximum
        } else {
            self.minimum
        }
    }
}

fn scale_product(left: f64, right: f64, factor: f64) -> Result<f64, ScalarError> {
    if left == 0.0 || right == 0.0 || factor == 0.0 {
        return Ok(0.0);
    }
    if !left.is_finite() || !right.is_finite() || !factor.is_finite() || factor < 0.0 {
        return Err(ScalarError::Number);
    }
    let (left_mantissa, left_exponent) = binary_parts(left.abs());
    let (right_mantissa, right_exponent) = binary_parts(right.abs());
    let (factor_mantissa, factor_exponent) = binary_parts(factor);
    let mut mantissa = left_mantissa * right_mantissa * factor_mantissa;
    let mut exponent = i64::from(left_exponent)
        .checked_add(i64::from(right_exponent))
        .and_then(|value| value.checked_add(i64::from(factor_exponent)))
        .ok_or(ScalarError::Number)?;
    while mantissa >= 2.0 {
        mantissa *= 0.5;
        exponent = exponent.checked_add(1).ok_or(ScalarError::Number)?;
    }
    while mantissa < 1.0 {
        mantissa *= 2.0;
        exponent = exponent.checked_sub(1).ok_or(ScalarError::Number)?;
    }
    let result = scale_binary(mantissa, exponent);
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ScalarError::Number)
    }
}
