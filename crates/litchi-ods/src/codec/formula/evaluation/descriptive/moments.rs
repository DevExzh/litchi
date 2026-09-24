//! Exact, bounded arithmetic for centered descriptive statistics.
//!
//! The ordinary `f64` mean/deviation path is attractive for large ranges, but
//! it cannot distinguish an exact cancellation from a small rounded value. In
//! particular, subtracting a rounded mean turns adjacent large numbers into
//! the wrong variance and can erase a small skew. This module keeps the four
//! raw power sums as signed binary integers. Each sum has a fixed quantum,
//! chosen from the smallest binary64 exponent, so every admitted binary64 is
//! represented exactly and no input cell is retained.
//!
//! The state is intentionally fixed-size. A binary64 has at most 53
//! significand bits and an exponent in `[-1074, 971]`; the four power sums
//! therefore fit comfortably in `STATE_LIMBS`, including the `u64` count.
//! Formula products use a larger fixed scratch width. An exhausted width is
//! treated as an arithmetic failure, rather than silently dropping a tail.
//! With the binary64 bounds and the checked count this path does not exhaust
//! the selected widths for a valid input sequence.

use super::super::dyadic::{
    self, ArithmeticError, Big, SignedBig, SignedDyadic, decompose,
    normalize_mantissa as normalize_dyadic_mantissa, round_ratio, round_signed_ratio,
    scale_power_of_two,
};

/// Raw sums and AVEDEV replay use this many 64-bit limbs.
const STATE_LIMBS: usize = 160;

/// Formula products and exact quotient rounding use this many limbs.
const CALC_LIMBS: usize = 384;

/// The fixed binary exponent of a represented binary64 significand.
const MIN_BINARY_EXPONENT: i64 = dyadic::MIN_BINARY_EXPONENT;

impl From<ArithmeticError> for MomentError {
    fn from(_: ArithmeticError) -> Self {
        Self::Number
    }
}

/// Failure classes understood by the resolver/scalar adapter.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum MomentError {
    /// No admitted number was available for AVEDEV or DEVSQ.
    Empty,
    /// The function's minimum numeric count was not met.
    MinimumCount,
    /// The admitted values have exactly zero variance.
    ZeroVariance,
    /// A non-finite value or an impossible bounded arithmetic state occurred.
    Number,
}

/// Highest raw power retained by an [`ExactMoments`] state.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
#[repr(u8)]
pub(super) enum MomentOrder {
    /// Retain only `S1` (AVEDEV).
    First = 1,
    /// Retain `S1` and `S2` (DEVSQ).
    Second = 2,
    /// Retain through `S3` (SKEW/SKEWP).
    Third = 3,
    /// Retain through `S4` (KURT).
    Fourth = 4,
}

impl MomentOrder {
    fn includes(self, power: u8) -> bool {
        self as u8 >= power
    }
}

/// Exact raw sums of powers one through four.
pub(super) struct ExactMoments {
    order: MomentOrder,
    count: u64,
    first: SignedDyadic<STATE_LIMBS>,
    second: SignedDyadic<STATE_LIMBS>,
    third: SignedDyadic<STATE_LIMBS>,
    fourth: SignedDyadic<STATE_LIMBS>,
}

impl ExactMoments {
    /// Create an empty fixed-size state.
    #[cfg(test)]
    pub(super) fn new() -> Self {
        Self::with_order(MomentOrder::Fourth)
    }

    /// Create a state retaining only the raw powers required by one reducer.
    pub(super) fn with_order(order: MomentOrder) -> Self {
        Self {
            order,
            count: 0,
            first: SignedDyadic::zero(),
            second: SignedDyadic::zero(),
            third: SignedDyadic::zero(),
            fourth: SignedDyadic::zero(),
        }
    }

    /// Add one finite binary64 exactly to all four raw sums.
    pub(super) fn push(&mut self, value: f64) -> Result<(), MomentError> {
        let (negative, mantissa, exponent) = decompose(value).ok_or(MomentError::Number)?;
        let next_count = self.count.checked_add(1).ok_or(MomentError::Number)?;

        // The mantissa products are at most 212 bits.  `add_power` places
        // those occupied limbs directly at the common quantum instead of
        // shifting a full fixed-width temporary for every cell.
        self.first.add_power(negative, mantissa, exponent, 1)?;
        if self.order.includes(2) {
            self.second.add_power(false, mantissa, exponent, 2)?;
        }
        if self.order.includes(3) {
            self.third.add_power(negative, mantissa, exponent, 3)?;
        }
        if self.order.includes(4) {
            self.fourth.add_power(false, mantissa, exponent, 4)?;
        }
        self.count = next_count;
        Ok(())
    }

    /// Prepare the exact second pass used by AVEDEV.
    pub(super) fn ave_dev_plan(&self) -> Result<AveDevPlan, MomentError> {
        if self.count == 0 {
            return Err(MomentError::Empty);
        }
        let sum = to_calc_signed(&self.first);
        Ok(AveDevPlan {
            count: self.count,
            sum,
        })
    }

    /// Return exact DEVSQ, rounded once to binary64.
    pub(super) fn finish_devsq(&self) -> Result<f64, MomentError> {
        if self.count == 0 {
            return Err(MomentError::Empty);
        }
        if !self.order.includes(2) {
            return Err(MomentError::Number);
        }
        let centered = self.centered(MomentOrder::Second)?;
        if centered.second.is_zero() {
            return Ok(0.0);
        }
        let denominator = Big::<CALC_LIMBS>::from_small(self.count);
        Ok(round_ratio(&centered.second, &denominator, -2148, false)?)
    }

    /// Return population (`sample == false`) or sample (`sample == true`)
    /// skew. The central numerator is exact; only the final standardized
    /// ratio is rounded through the normalized binary path.
    pub(super) fn finish_skew(&self, sample: bool) -> Result<f64, MomentError> {
        if self.count < 3 {
            return Err(MomentError::MinimumCount);
        }
        if !self.order.includes(3) {
            return Err(MomentError::Number);
        }
        let centered = self.centered(MomentOrder::Third)?;
        if centered.second.is_zero() {
            return Err(MomentError::ZeroVariance);
        }
        let third = centered.third.as_ref().ok_or(MomentError::Number)?;
        let mut result = standardized_third(third, &centered.second)?;
        if sample {
            let n = self.count as f64;
            let factor = (n * (n - 1.0)).sqrt() / (n - 2.0);
            result.mantissa *= factor;
            normalize_dyadic_mantissa(&mut result.mantissa, &mut result.exponent)?;
        }
        let result = result.value();
        if !result.is_finite() {
            return Err(MomentError::Number);
        }
        Ok(result)
    }

    /// Return exact sample excess kurtosis, including the cancellation in the
    /// numerator, rounded only after the complete rational is formed.
    pub(super) fn finish_kurt(&self) -> Result<f64, MomentError> {
        if self.count < 4 {
            return Err(MomentError::MinimumCount);
        }
        if !self.order.includes(4) {
            return Err(MomentError::Number);
        }
        let centered = self.centered(MomentOrder::Fourth)?;
        if centered.second.is_zero() {
            return Err(MomentError::ZeroVariance);
        }

        let square = multiply(&centered.second, &centered.second)?;
        let n = self.count;
        let n_plus_one = n.checked_add(1).ok_or(MomentError::Number)?;
        let n_minus_one = n - 1;
        let n_minus_two = n - 2;
        let n_minus_three = n - 3;

        let c4 = centered.fourth.as_ref().ok_or(MomentError::Number)?;
        let mut c4_term = *c4;
        c4_term.mul_small(n_plus_one)?;
        let mut c2_term = square;
        c2_term.mul_small(3)?;
        c2_term.mul_small(n_minus_one)?;
        let mut numerator = SignedBig::from_positive(&c4_term);
        numerator.sub_unsigned(&c2_term)?;
        numerator.mul_small(n_minus_one)?;

        let mut denominator = square;
        denominator.mul_small(n_minus_two)?;
        denominator.mul_small(n_minus_three)?;
        if numerator.is_zero() {
            return Ok(0.0);
        }
        Ok(round_signed_ratio(&numerator, &denominator, 0)?)
    }

    fn centered(&self, need: MomentOrder) -> Result<CenteredNumerators, MomentError> {
        if !self.order.includes(need as u8) {
            return Err(MomentError::Number);
        }
        let first = to_calc_signed(&self.first);
        let second = to_calc_signed(&self.second);
        let n = self.count;

        // C2 = n*S2 - S1².
        let first_square = multiply(&first.magnitude, &first.magnitude)?;
        let mut second_term = second.magnitude;
        second_term.mul_small(n)?;
        let mut c2 = SignedBig::from_positive(&second_term);
        c2.sub_unsigned(&first_square)?;
        if c2.negative {
            return Err(MomentError::Number);
        }

        let third = if need.includes(3) {
            let third = to_calc_signed(&self.third);
            // C3 = n²*S3 - 3*n*S1*S2 + 2*S1³.
            let mut c3 = SignedBig::zero();
            let mut term = third.magnitude;
            term.mul_small(n)?;
            term.mul_small(n)?;
            c3.add_unsigned(third.negative, &term)?;

            let first_second = multiply(&first.magnitude, &second.magnitude)?;
            let mut term = first_second;
            term.mul_small(3)?;
            term.mul_small(n)?;
            c3.add_unsigned(!first.negative, &term)?;

            let mut first_cube = multiply(&first.magnitude, &first.magnitude)?;
            first_cube = multiply(&first_cube, &first.magnitude)?;
            first_cube.mul_small(2)?;
            c3.add_unsigned(first.negative, &first_cube)?;
            Some(c3)
        } else {
            None
        };

        let fourth = if need.includes(4) {
            let third = to_calc_signed(&self.third);
            let fourth = to_calc_signed(&self.fourth);
            // C4 = n³*S4 - 4*n²*S1*S3 + 6*n*S1²*S2 - 3*S1⁴.
            let mut c4 = SignedBig::zero();
            let mut term = fourth.magnitude;
            term.mul_small(n)?;
            term.mul_small(n)?;
            term.mul_small(n)?;
            c4.add_unsigned(fourth.negative, &term)?;

            let first_third = multiply(&first.magnitude, &third.magnitude)?;
            let mut term = first_third;
            term.mul_small(4)?;
            term.mul_small(n)?;
            term.mul_small(n)?;
            c4.add_unsigned(!(first.negative ^ third.negative), &term)?;

            let first_square_second = multiply(&first.magnitude, &first.magnitude)?;
            let mut first_square_second = multiply(&first_square_second, &second.magnitude)?;
            first_square_second.mul_small(6)?;
            first_square_second.mul_small(n)?;
            c4.add_unsigned(false, &first_square_second)?;

            let mut first_fourth = multiply(&first.magnitude, &first.magnitude)?;
            first_fourth = multiply(&first_fourth, &first.magnitude)?;
            first_fourth = multiply(&first_fourth, &first.magnitude)?;
            first_fourth.mul_small(3)?;
            c4.add_unsigned(true, &first_fourth)?;

            if c4.negative {
                return Err(MomentError::Number);
            }
            Some(c4.magnitude)
        } else {
            None
        };
        Ok(CenteredNumerators {
            second: c2.magnitude,
            third,
            fourth,
        })
    }
}

/// The first-pass state needed by AVEDEV's exact absolute-deviation replay.
#[derive(Clone)]
pub(super) struct AveDevPlan {
    count: u64,
    sum: SignedBig<CALC_LIMBS>,
}

impl AveDevPlan {
    /// Start a replay over the same admitted sequence.
    pub(super) fn replay(&self) -> AveDevReplay {
        AveDevReplay {
            expected: self.count,
            seen: 0,
            sum: Big::zero(),
            mean_sum: self.sum,
        }
    }
}

/// Exact absolute deviations for AVEDEV.
pub(super) struct AveDevReplay {
    expected: u64,
    seen: u64,
    sum: Big<CALC_LIMBS>,
    mean_sum: SignedBig<CALC_LIMBS>,
}

impl AveDevReplay {
    /// Add `|n*x - S1|` at the common `2^-1074` quantum.
    pub(super) fn push(&mut self, value: f64) -> Result<(), MomentError> {
        if self.seen >= self.expected {
            return Err(MomentError::Number);
        }
        let (negative, mantissa, exponent) = decompose(value).ok_or(MomentError::Number)?;
        let mut term = Big::<CALC_LIMBS>::from_small(mantissa);
        term.mul_small(self.expected)?;
        let shift =
            usize::try_from(exponent - MIN_BINARY_EXPONENT).map_err(|_| MomentError::Number)?;
        term.shift_left(shift)?;
        let mut delta = SignedBig::from_unsigned(negative, term);
        delta.sub_signed(&self.mean_sum)?;
        self.sum.add_assign(&delta.magnitude)?;
        self.seen = self.seen.checked_add(1).ok_or(MomentError::Number)?;
        Ok(())
    }

    /// Finish `sum(|n*x-S1|)/n²` with exact binary rounding.
    pub(super) fn finish(&self) -> Result<f64, MomentError> {
        if self.seen != self.expected {
            return Err(MomentError::Number);
        }
        if self.sum.is_zero() {
            return Ok(0.0);
        }
        let denominator = Big::<CALC_LIMBS>::from_small(self.expected);
        let mut denominator = denominator;
        denominator.mul_small(self.expected)?;
        Ok(round_ratio(
            &self.sum,
            &denominator,
            MIN_BINARY_EXPONENT,
            false,
        )?)
    }
}

struct CenteredNumerators {
    second: Big<CALC_LIMBS>,
    third: Option<SignedBig<CALC_LIMBS>>,
    fourth: Option<Big<CALC_LIMBS>>,
}

fn to_calc_signed<const N: usize>(value: &SignedDyadic<N>) -> SignedBig<CALC_LIMBS> {
    let mut magnitude = Big::<CALC_LIMBS>::zero();
    for index in 0..value.magnitude.len {
        magnitude.limbs[index] = value.magnitude.limbs[index];
    }
    magnitude.len = value.magnitude.len;
    SignedBig::from_unsigned(value.negative, magnitude)
}

fn multiply<const N: usize>(left: &Big<N>, right: &Big<N>) -> Result<Big<N>, MomentError> {
    dyadic::multiply(left, right).map_err(Into::into)
}

/// A normalized binary value retained before final scaling. Sample skew
/// applies its correction to this representation so a tiny value is rounded
/// only once.
#[derive(Clone, Copy)]
struct NormalizedBinary {
    negative: bool,
    mantissa: f64,
    exponent: i64,
}

impl NormalizedBinary {
    fn value(self) -> f64 {
        let magnitude = scale_power_of_two(self.mantissa, self.exponent);
        if self.negative { -magnitude } else { magnitude }
    }
}

fn standardized_third(
    third: &SignedBig<CALC_LIMBS>,
    second: &Big<CALC_LIMBS>,
) -> Result<NormalizedBinary, MomentError> {
    if third.is_zero() {
        return Ok(NormalizedBinary {
            negative: false,
            mantissa: 0.0,
            exponent: 0,
        });
    }
    let (third_mantissa, third_exponent) = third.magnitude.normalized();
    let (second_mantissa, second_exponent) = second.normalized();
    let denominator = second_mantissa * second_mantissa.sqrt();
    let mut quotient = third_mantissa / denominator;
    let exponent_twice = third_exponent
        .checked_mul(2)
        .and_then(|value| {
            second_exponent
                .checked_mul(3)
                .and_then(|other| value.checked_sub(other))
        })
        .ok_or(MomentError::Number)?;
    let mut exponent = exponent_twice.div_euclid(2);
    if exponent_twice.rem_euclid(2) != 0 {
        quotient *= 2.0_f64.sqrt();
    }
    normalize_dyadic_mantissa(&mut quotient, &mut exponent)?;
    Ok(NormalizedBinary {
        negative: third.negative,
        mantissa: quotient,
        exponent,
    })
}

#[cfg(test)]
mod tests {
    use super::{Big, ExactMoments, MomentError, MomentOrder};

    fn assert_close(actual: f64, expected: f64) {
        let scale = expected.abs().max(1.0);
        assert!(
            (actual - expected).abs() <= 8.0 * f64::EPSILON * scale,
            "actual={actual:?} expected={expected:?}"
        );
    }

    #[test]
    fn exact_basic_centered_reducers() {
        let mut moments = ExactMoments::new();
        for value in [1.0, 2.0, 3.0] {
            moments.push(value).unwrap();
        }
        assert_eq!(moments.finish_devsq().unwrap(), 2.0);
        assert_eq!(moments.finish_skew(false).unwrap(), 0.0);
        assert_eq!(moments.finish_skew(true).unwrap(), 0.0);
        let plan = moments.ave_dev_plan().unwrap();
        let mut replay = plan.replay();
        for value in [1.0, 2.0, 3.0] {
            replay.push(value).unwrap();
        }
        assert_close(replay.finish().unwrap(), 2.0 / 3.0);
    }

    #[test]
    fn adjacent_large_values_keep_half_ulp_center() {
        let mut moments = ExactMoments::new();
        let first = 4_503_599_627_370_496.0; // 2^52
        moments.push(first).unwrap();
        moments.push(first + 1.0).unwrap();
        assert_eq!(moments.finish_devsq().unwrap(), 0.5);
    }

    #[test]
    fn exact_cancellation_is_zero_skew() {
        let mut moments = ExactMoments::new();
        for value in [-1.0, 0.0, 1.0] {
            moments.push(value).unwrap();
        }
        assert_eq!(moments.finish_devsq().unwrap(), 2.0);
        assert_eq!(moments.finish_skew(false).unwrap(), 0.0);
    }

    #[test]
    fn direct_subnormal_rounding_keeps_both_sides_of_half_minimum() {
        use super::super::super::dyadic::{Big, round_ratio};

        let denominator = Big::<8>::from_small(1_u64 << 56);
        let above = Big::<8>::from_small((1_u64 << 55) + 1);
        let tie = Big::<8>::from_small(1_u64 << 55);
        let below = Big::<8>::from_small((1_u64 << 55) - 1);
        assert_eq!(
            round_ratio(&above, &denominator, -1074, false)
                .unwrap()
                .to_bits(),
            1
        );
        assert_eq!(
            round_ratio(&tie, &denominator, -1074, false)
                .unwrap()
                .to_bits(),
            0
        );
        assert_eq!(
            round_ratio(&below, &denominator, -1074, false)
                .unwrap()
                .to_bits(),
            0
        );
        assert_eq!(
            round_ratio(&above, &denominator, -1074, true)
                .unwrap()
                .to_bits(),
            1_u64 << 63 | 1
        );
    }

    #[test]
    fn direct_subnormal_rounding_preserves_low_remainder() {
        use super::super::super::dyadic::{Big, round_ratio};

        let numerator = Big::<8>::from_small((1_u64 << 54) + 11);
        let denominator = Big::<8>::from_small(1);
        assert_eq!(
            round_ratio(&numerator, &denominator, -1077, false)
                .unwrap()
                .to_bits(),
            0x0008_0000_0000_0001
        );
    }

    #[test]
    fn kurtosis_forms_complete_rational_before_rounding() {
        let mut moments = ExactMoments::new();
        for value in [1.0, 2.0, 3.0, 4.0] {
            moments.push(value).unwrap();
        }
        assert_close(moments.finish_kurt().unwrap(), -1.2);
    }

    #[test]
    fn minimum_and_zero_variance_domains_are_distinct() {
        let mut one = ExactMoments::new();
        one.push(1.0).unwrap();
        assert_eq!(one.finish_devsq().unwrap(), 0.0);
        assert_eq!(one.finish_skew(false), Err(MomentError::MinimumCount));
        assert_eq!(one.finish_kurt(), Err(MomentError::MinimumCount));

        let mut equal = ExactMoments::new();
        equal.push(4.0).unwrap();
        equal.push(4.0).unwrap();
        equal.push(4.0).unwrap();
        assert_eq!(equal.finish_skew(false), Err(MomentError::ZeroVariance));
    }

    #[test]
    fn specialized_orders_skip_unneeded_raw_powers() {
        let values = [1.0, 2.0, 3.0, 4.0];

        let mut first = ExactMoments::with_order(MomentOrder::First);
        for value in values {
            first.push(value).unwrap();
        }
        let plan = first.ave_dev_plan().unwrap();
        let mut replay = plan.replay();
        for value in values {
            replay.push(value).unwrap();
        }
        assert_close(replay.finish().unwrap(), 1.0);
        assert_eq!(first.finish_devsq(), Err(MomentError::Number));

        let mut second = ExactMoments::with_order(MomentOrder::Second);
        for value in values {
            second.push(value).unwrap();
        }
        assert_eq!(second.finish_devsq().unwrap(), 5.0);
        assert_eq!(second.finish_skew(false), Err(MomentError::Number));

        let mut third = ExactMoments::with_order(MomentOrder::Third);
        for value in values {
            third.push(value).unwrap();
        }
        assert!(third.finish_skew(false).unwrap().abs() < 1e-12);
        assert_eq!(third.finish_kurt(), Err(MomentError::Number));
    }

    #[test]
    fn clearing_a_big_integer_does_not_resurrect_stale_limbs() {
        let mut value = Big::<4>::from_small(7);
        value.clear();
        value.add_assign(&Big::<4>::from_small(3)).unwrap();
        assert_eq!(value.len, 1);
        assert_eq!(value.limbs[0], 3);
    }

    #[test]
    fn wide_signed_ave_dev_is_order_independent() {
        let cases = [
            [-1.0e16, 2.0, 3.0, 1.0e16, 1.0],
            [1.0, 1.0e16, 3.0, 2.0, -1.0e16],
            [3.0, -1.0e16, 1.0, 1.0e16, 2.0],
            [1.0e16, 2.0, -1.0e16, 1.0, 3.0],
        ];
        let expected = 4_000_000_000_000_000.5;
        for values in cases {
            let mut moments = ExactMoments::new();
            for value in values {
                moments.push(value).unwrap();
            }
            let plan = moments.ave_dev_plan().unwrap();
            let mut replay = plan.replay();
            for value in values {
                replay.push(value).unwrap();
            }
            assert_eq!(replay.finish().unwrap(), expected);
        }
    }

    #[test]
    fn shifted_subtraction_propagates_past_zero_high_carry() {
        let mut value = Big::<4>::zero();
        value.limbs[2] = 1;
        value.len = 3;
        let term = Big::<4>::from_small(1);
        value.sub_shifted_assign(&term, 65).unwrap();
        assert_eq!(value.len, 2);
        assert_eq!(value.limbs[0], 0);
        assert_eq!(value.limbs[1], u64::MAX - 1);
    }

    #[test]
    fn skew_keeps_exact_signed_power_cancellation_at_shift_boundary() {
        let values = [
            -f64::from_bits(0x0810_0000_0000_0000),
            f64::from_bits(0x0420_0000_0000_0000),
            f64::from_bits(0x0170_0000_0000_0000),
        ];
        let mut moments = ExactMoments::new();
        for value in values {
            moments.push(value).unwrap();
        }
        assert_eq!(
            moments.finish_skew(false).unwrap().to_bits(),
            0xbfe6_a09e_667f_3bcd
        );
        assert_eq!(
            moments.finish_skew(true).unwrap().to_bits(),
            0xbffb_b67a_e858_4caa
        );
    }
}
