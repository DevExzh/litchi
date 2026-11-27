//! Fixed-width first pass for signed sums of binary64 reciprocals.
//!
//! `HARMEAN` cannot safely form `1.0 / x` for every finite input: a
//! reciprocal can overflow for a subnormal input or underflow for a value
//! near `f64::MAX`.  This accumulator keeps the rounded reciprocal
//! significands as integers at one fixed binary quantum instead.  The three
//! magnitudes are fixed arrays, so a reference scan never retains its input
//! cells and never allocates while it is adding a value.
//!
//! For an input represented as `sign * m * 2^e`, where `m` is odd, let
//! `k=floor(log2(m))` and
//!
//! ```text
//! Q = round_even(2^(k+53) / m)
//! 1/x ~= sign * Q * 2^(-e-k-53)
//! ```
//!
//! `Q` is at most 54 bits including the exact `2^53` endpoint.  A non-exact
//! division contributes an outward error of at most `2^(-e-k-54)`.  We store
//! both the rounded term and that error in
//! units of `2^-1100`.  The full binary64 exponent span plus a `u64` number of
//! terms fits in 36 little-endian `u64` limbs.  At finish the positive and
//! negative magnitudes are subtracted exactly; cancellation therefore causes
//! a replay request instead of being hidden by a floating-point tolerance.

const LIMBS: usize = 36;
const BASE_EXPONENT: i64 = -1100;
const SIGNIFICAND_BITS: u32 = 53;

/// A fixed-width signed reciprocal sum and its outward rounding bound.
///
/// This type is intentionally `Copy`: the state is bounded and contains no
/// pointers or allocation.  The evaluator owns the admitted-value count and
/// passes it to [`Self::try_finish`], which keeps the accumulator useful for
/// both scalar and resolver-backed reducers.
#[derive(Clone, Copy, Debug)]
pub(super) struct ReciprocalSum {
    positive: [u64; LIMBS],
    negative: [u64; LIMBS],
    error: [u64; LIMBS],
    invalid: bool,
    overflow: bool,
}

impl ReciprocalSum {
    /// Create an empty fixed-width reciprocal sum.
    pub(super) const fn new() -> Self {
        Self {
            positive: [0; LIMBS],
            negative: [0; LIMBS],
            error: [0; LIMBS],
            invalid: false,
            overflow: false,
        }
    }

    /// Add one finite, non-zero binary64 value.
    ///
    /// The harmonic reducer validates its domain before calling this method.
    /// Invalid values are retained as a failed first pass so an accidental
    /// caller cannot turn a zero or non-finite input into a plausible result.
    pub(super) fn push(&mut self, value: f64) {
        let Some((negative, mantissa, exponent)) = decompose(value) else {
            self.invalid = true;
            return;
        };

        let k = u64::BITS - mantissa.leading_zeros() - 1;
        let numerator = 1_u128 << (k + SIGNIFICAND_BITS);
        let divisor = u128::from(mantissa);
        let quotient = numerator / divisor;
        let remainder = numerator % divisor;
        let rounded = round_even(quotient, remainder, divisor);

        // The primary term is Q * 2^(-e-k-53).  The error quantum is one
        // half-unit at that scale.  All exponents are checked against the
        // fixed representation even though the binary64 derivation proves
        // they fit; a failed proof must request replay rather than truncate.
        let term_exponent = 0_i64
            .checked_sub(exponent)
            .and_then(|value| value.checked_sub(i64::from(k)))
            .and_then(|value| value.checked_sub(i64::from(SIGNIFICAND_BITS)))
            .and_then(|value| value.checked_sub(BASE_EXPONENT));
        let Some(term_shift) = term_exponent.and_then(nonnegative_shift) else {
            self.overflow = true;
            return;
        };
        if !add_small(
            if negative {
                &mut self.negative
            } else {
                &mut self.positive
            },
            rounded as u64,
            term_shift,
        ) {
            self.overflow = true;
            return;
        }

        if remainder != 0 {
            let Some(error_shift) = term_shift.checked_sub(1) else {
                self.overflow = true;
                return;
            };
            if !add_bit(&mut self.error, error_shift) {
                self.overflow = true;
            }
        }
    }

    /// Return a certified finite harmonic denominator quotient.
    ///
    /// `None` means that exact/adaptive replay is required.  This includes
    /// exact signed cancellation, an uncertain sign, an accumulator bound
    /// failure, and an output that cannot be certified as a non-zero finite
    /// binary64.  The caller maps an exact replay zero denominator to its
    /// formula-level division-by-zero error.
    pub(super) fn try_finish(&self, count: u64) -> Option<f64> {
        if count == 0 || self.invalid || self.overflow {
            return None;
        }

        let (sum, negative) = subtract_magnitudes(&self.positive, &self.negative)?;
        if !bound_is_small(&self.error, &sum) {
            return None;
        }

        // `bound_is_small` proves an input-denominator relative error of at
        // most 2^-52.  The rounded 53-bit denominator, the u64-to-binary64
        // count conversion, and the quotient operation each add at most one
        // half-ulp-scale relative error; final binary scaling contributes
        // one correctly-rounded ulp.  The resulting bound is below eight
        // ulps.  A result that rounds to zero or overflows is deliberately
        // left to replay by `scale_binary`, where this relative estimate
        // would not be a useful absolute certificate.
        let (mantissa, exponent) = rounded_mantissa(&sum)?;
        let count_mantissa = count as f64;
        if !count_mantissa.is_finite() || count_mantissa == 0.0 {
            return None;
        }

        // The exact rounded sum is approximately
        // `mantissa * 2^(exponent + BASE_EXPONENT)`.  Keep the ratio in
        // [1,2) before applying the binary exponent so neither an enormous
        // reciprocal nor an underflowing reciprocal is formed as f64.
        let mut ratio = count_mantissa / mantissa;
        if !ratio.is_finite() || ratio == 0.0 {
            return None;
        }
        let mut result_exponent = BASE_EXPONENT
            .checked_neg()
            .and_then(|value| value.checked_add(52))
            .and_then(|value| value.checked_sub(exponent))?;
        while ratio >= 2.0 {
            ratio *= 0.5;
            result_exponent = result_exponent.checked_add(1)?;
        }
        while ratio < 1.0 {
            ratio *= 2.0;
            result_exponent = result_exponent.checked_sub(1)?;
        }
        let result = scale_binary(ratio, result_exponent)?;
        if result == 0.0 || !result.is_finite() {
            None
        } else if negative {
            Some(-result)
        } else {
            Some(result)
        }
    }
}

impl Default for ReciprocalSum {
    fn default() -> Self {
        Self::new()
    }
}

/// Decompose a finite non-zero binary64 as `sign * odd_mantissa * 2^e`.
fn decompose(value: f64) -> Option<(bool, u64, i64)> {
    if !value.is_finite() || value == 0.0 {
        return None;
    }
    let bits = value.to_bits();
    let negative = bits >> 63 != 0;
    let raw_exponent = (bits >> 52) & 0x7ff;
    let fraction = bits & ((1_u64 << 52) - 1);
    let (significand, exponent) = if raw_exponent == 0 {
        (fraction, -1074_i64)
    } else {
        ((1_u64 << 52) | fraction, raw_exponent as i64 - 1023 - 52)
    };
    if significand == 0 {
        return None;
    }
    let trailing = i64::from(significand.trailing_zeros());
    Some((negative, significand >> trailing, exponent + trailing))
}

fn round_even(quotient: u128, remainder: u128, divisor: u128) -> u128 {
    let doubled = remainder << 1;
    if doubled > divisor || (doubled == divisor && quotient & 1 != 0) {
        quotient + 1
    } else {
        quotient
    }
}

fn nonnegative_shift(value: i64) -> Option<usize> {
    usize::try_from(value).ok()
}

/// Add a 54-bit-or-smaller term at a bit offset, touching only its two source
/// limbs and any carry chain.  The chain is bounded by the fixed array and is
/// not a full-array scan for ordinary additions.
fn add_small(limbs: &mut [u64; LIMBS], value: u64, shift: usize) -> bool {
    if value == 0 {
        return true;
    }
    let word = shift / 64;
    let bit = shift % 64;
    if word >= LIMBS {
        return false;
    }

    let low = (value as u128) << bit;
    let low_limb = low as u64;
    let high_limb = (low >> 64) as u64;
    if !add_limb(limbs, word, low_limb) {
        return false;
    }
    if high_limb != 0 && !add_limb(limbs, word + 1, high_limb) {
        return false;
    }
    true
}

fn add_bit(limbs: &mut [u64; LIMBS], shift: usize) -> bool {
    let word = shift / 64;
    let bit = shift % 64;
    if word >= LIMBS {
        return false;
    }
    add_limb(limbs, word, 1_u64 << bit)
}

fn add_limb(limbs: &mut [u64; LIMBS], mut index: usize, value: u64) -> bool {
    if index >= LIMBS {
        return false;
    }
    let (sum, mut carry) = limbs[index].overflowing_add(value);
    limbs[index] = sum;
    index += 1;
    while carry {
        if index >= LIMBS {
            return false;
        }
        let (next, next_carry) = limbs[index].overflowing_add(1);
        limbs[index] = next;
        carry = next_carry;
        index += 1;
    }
    true
}

/// Subtract the two magnitudes exactly, returning the larger sign and a copy
/// of the absolute difference.
fn subtract_magnitudes(
    positive: &[u64; LIMBS],
    negative: &[u64; LIMBS],
) -> Option<([u64; LIMBS], bool)> {
    match compare_magnitudes(positive, negative) {
        core::cmp::Ordering::Equal => None,
        core::cmp::Ordering::Greater => Some((subtract(positive, negative), false)),
        core::cmp::Ordering::Less => Some((subtract(negative, positive), true)),
    }
}

fn compare_magnitudes(left: &[u64; LIMBS], right: &[u64; LIMBS]) -> core::cmp::Ordering {
    for index in (0..LIMBS).rev() {
        match left[index].cmp(&right[index]) {
            core::cmp::Ordering::Equal => {},
            ordering => return ordering,
        }
    }
    core::cmp::Ordering::Equal
}

fn subtract(left: &[u64; LIMBS], right: &[u64; LIMBS]) -> [u64; LIMBS] {
    let mut result = [0_u64; LIMBS];
    let mut borrow = false;
    for index in 0..LIMBS {
        let (first, first_borrow) = left[index].overflowing_sub(right[index]);
        let (value, second_borrow) = first.overflowing_sub(u64::from(borrow));
        result[index] = value;
        borrow = first_borrow || second_borrow;
    }
    debug_assert!(!borrow);
    result
}

/// Check `error * 2^52 <= sum` without using a floating-point tolerance.
fn bound_is_small(error: &[u64; LIMBS], sum: &[u64; LIMBS]) -> bool {
    let mut shifted = [0_u64; LIMBS];
    let bit_shift = 52_u32;
    if error[LIMBS - 1] >> (64 - bit_shift) != 0 {
        return false;
    }
    for index in 0..LIMBS {
        shifted[index] |= error[index] << bit_shift;
        if index > 0 {
            shifted[index] |= error[index - 1] >> (64 - bit_shift);
        }
    }
    compare_magnitudes(&shifted, sum) != core::cmp::Ordering::Greater
}

/// Round a positive fixed-point integer to a 53-bit significand and return
/// `(mantissa, exponent)` for `sum ~= mantissa * 2^(exponent-52)`.
fn rounded_mantissa(sum: &[u64; LIMBS]) -> Option<(f64, i64)> {
    let highest = highest_bit(sum)?;
    let (mantissa_bits, exponent) = if highest <= 52 {
        let shift = 52 - highest;
        (shift_left_u64(sum, shift)?, i64::try_from(highest).ok()?)
    } else {
        let shift = highest - 52;
        (rounded_top_bits(sum, shift)?, i64::try_from(highest).ok()?)
    };
    let mut mantissa = mantissa_bits;
    let mut exponent = exponent;
    if mantissa >= (1_u64 << 53) {
        mantissa >>= 1;
        exponent = exponent.checked_add(1)?;
    }
    if mantissa < (1_u64 << 52) {
        return None;
    }
    Some((mantissa as f64, exponent))
}

fn highest_bit(limbs: &[u64; LIMBS]) -> Option<usize> {
    for index in (0..LIMBS).rev() {
        let value = limbs[index];
        if value != 0 {
            return Some(index * 64 + (63 - value.leading_zeros() as usize));
        }
    }
    None
}

fn shift_left_u64(sum: &[u64; LIMBS], shift: usize) -> Option<u64> {
    debug_assert!(shift <= 52);
    let mut value = 0_u64;
    for bit in 0_usize..=52 {
        if let Some(source) = bit.checked_sub(shift) {
            if source < LIMBS * 64 && bit_is_set(sum, source) {
                value |= 1_u64 << bit;
            }
        }
    }
    Some(value)
}

fn rounded_top_bits(sum: &[u64; LIMBS], shift: usize) -> Option<u64> {
    let mut value = 0_u64;
    for bit in 0..53 {
        let source = shift.checked_add(bit)?;
        if bit_is_set(sum, source) {
            value |= 1_u64 << bit;
        }
    }
    let guard = shift.checked_sub(1).is_some_and(|bit| bit_is_set(sum, bit));
    let sticky = shift
        .checked_sub(1)
        .is_some_and(|last| (0..last).any(|bit| bit_is_set(sum, bit)));
    if guard && (sticky || value & 1 != 0) {
        value = value.checked_add(1)?;
    }
    Some(value)
}

fn bit_is_set(limbs: &[u64; LIMBS], bit: usize) -> bool {
    let word = bit / 64;
    word < LIMBS && (limbs[word] & (1_u64 << (bit % 64))) != 0
}

/// Scale a normalized `[1,2)` mantissa by a binary exponent with ties-to-even
/// subnormal rounding.  An overflow or a zero result requests replay.
fn scale_binary(value: f64, exponent: i64) -> Option<f64> {
    debug_assert!(value.is_finite() && (1.0..2.0).contains(&value));
    if exponent > 1023 {
        return None;
    }
    let bits = value.to_bits();
    let sign = bits & (1_u64 << 63);
    let fraction = bits & ((1_u64 << 52) - 1);
    if exponent >= -1022 {
        let target = u64::try_from(exponent + 1023).ok()?;
        return Some(f64::from_bits(sign | (target << 52) | fraction));
    }

    let shift = (-1022_i64).checked_sub(exponent)?;
    if shift >= 64 {
        return None;
    }
    let shift = u32::try_from(shift).ok()?;
    let significand = (1_u64 << 52) | fraction;
    let mask = (1_u64 << shift) - 1;
    let mut subnormal = significand >> shift;
    let remainder = significand & mask;
    if shift != 0 {
        let halfway = 1_u64 << (shift - 1);
        if remainder > halfway || (remainder == halfway && subnormal & 1 != 0) {
            subnormal = subnormal.checked_add(1)?;
        }
    }
    if subnormal == 0 {
        return None;
    }
    if subnormal >= (1_u64 << 52) {
        return Some(f64::from_bits(sign | (1_u64 << 52)));
    }
    Some(f64::from_bits(sign | subnormal))
}

#[cfg(test)]
mod tests {
    use super::ReciprocalSum;

    fn finish(values: &[f64]) -> Option<f64> {
        let mut sum = ReciprocalSum::new();
        for &value in values {
            sum.push(value);
        }
        sum.try_finish(values.len() as u64)
    }

    #[test]
    fn dyadic_values_finish_without_replay() {
        let value = finish(&[1.0, 2.0, 4.0, 8.0]).expect("dyadic sum");
        assert_eq!(value, 4.0 / (1.0 + 0.5 + 0.25 + 0.125));
    }

    #[test]
    fn signed_non_cancelling_values_finish_without_replay() {
        let value = finish(&[-2.0, 4.0]).expect("signed sum");
        assert_eq!(value, -8.0);
    }

    #[test]
    fn ordinary_fractional_reciprocals_round_to_the_exact_result() {
        assert_eq!(finish(&[3.0, 6.0]), Some(4.0));
        assert_eq!(finish(&[3.0, 6.0, 9.0]), Some(4.909090909090909));
    }

    #[test]
    fn exact_rational_cancellation_requests_replay() {
        assert_eq!(finish(&[3.0, 6.0, -2.0]), None);
    }

    #[test]
    fn extreme_subnormal_and_large_value_remain_bounded() {
        let value = finish(&[f64::from_bits(1), -f64::from_bits(1), 1.0e100]);
        assert!(value.is_none() || (value.unwrap() / 3.0e100 - 1.0).abs() < 1.0e-14);
    }

    #[test]
    fn positive_extreme_value_finishes() {
        let value = finish(&[f64::MAX, f64::MAX]).expect("finite harmonic mean");
        // The rounded reciprocal interval straddles the final max-value
        // rounding boundary.  The exact replay owns this edge case.
        assert!((value - f64::MAX).abs() <= f64::EPSILON * f64::MAX * 8.0);
    }

    #[test]
    fn mixed_exponents_keep_the_nonzero_tail() {
        let value = finish(&[f64::from_bits(1), 1.0e100]);
        assert_eq!(value, Some(2.0 * f64::from_bits(1)));
    }
}
