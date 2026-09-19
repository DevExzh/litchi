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

use core::cmp::Ordering;

/// Raw sums and AVEDEV replay use this many 64-bit limbs.
const STATE_LIMBS: usize = 160;

/// Formula products and exact quotient rounding use this many limbs.
const CALC_LIMBS: usize = 384;

/// The fixed binary exponent of a represented binary64 significand.
const MIN_BINARY_EXPONENT: i64 = -1074;

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
        round_ratio(&centered.second, &denominator, -2148, false)
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
            normalize_mantissa(&mut result.mantissa, &mut result.exponent)?;
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
        let mut c4_term = c4.clone();
        c4_term.mul_small(n_plus_one)?;
        let mut c2_term = square.clone();
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
        round_signed_ratio(&numerator, &denominator, 0)
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
        let mut second_term = second.magnitude.clone();
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
            let mut term = third.magnitude.clone();
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
            let mut term = fourth.magnitude.clone();
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
            mean_sum: self.sum.clone(),
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
        round_ratio(&self.sum, &denominator, MIN_BINARY_EXPONENT, false)
    }
}

struct CenteredNumerators {
    second: Big<CALC_LIMBS>,
    third: Option<SignedBig<CALC_LIMBS>>,
    fourth: Option<Big<CALC_LIMBS>>,
}

#[derive(Clone)]
struct SignedDyadic<const N: usize> {
    negative: bool,
    magnitude: Big<N>,
}

impl<const N: usize> SignedDyadic<N> {
    fn zero() -> Self {
        Self {
            negative: false,
            magnitude: Big::zero(),
        }
    }

    fn add_power(
        &mut self,
        negative: bool,
        mantissa: u64,
        exponent: i64,
        power: usize,
    ) -> Result<(), MomentError> {
        if mantissa == 0 {
            return Ok(());
        }
        let mut term = Big::<N>::from_small(1);
        for _ in 0..power {
            term.mul_small(mantissa)?;
        }
        let shift = exponent
            .checked_sub(MIN_BINARY_EXPONENT)
            .and_then(|difference| difference.checked_mul(power as i64))
            .and_then(|difference| usize::try_from(difference).ok())
            .ok_or(MomentError::Number)?;
        let term_negative = power % 2 == 1 && negative;

        if self.magnitude.is_zero() {
            self.magnitude.assign_shifted(&term, shift)?;
            self.negative = term_negative;
            return Ok(());
        }
        if self.negative == term_negative {
            self.magnitude.add_shifted_assign(&term, shift)?;
            return Ok(());
        }
        match cmp_shifted(&self.magnitude, &term, shift) {
            Ordering::Greater => self.magnitude.sub_shifted_assign(&term, shift)?,
            Ordering::Equal => {
                self.magnitude.clear();
                self.negative = false;
            },
            Ordering::Less => {
                let old = self.magnitude.clone();
                self.magnitude.assign_shifted(&term, shift)?;
                self.magnitude.sub_assign(&old);
                self.negative = term_negative;
            },
        }
        Ok(())
    }
}

fn decompose(value: f64) -> Option<(bool, u64, i64)> {
    if !value.is_finite() {
        return None;
    }
    let bits = value.to_bits();
    let negative = (bits >> 63) != 0;
    let exponent_bits = ((bits >> 52) & 0x7ff) as u16;
    let fraction = bits & ((1_u64 << 52) - 1);
    let (mantissa, exponent) = if exponent_bits == 0 {
        if fraction == 0 {
            return Some((negative, 0, MIN_BINARY_EXPONENT));
        }
        (fraction, MIN_BINARY_EXPONENT)
    } else {
        (
            (1_u64 << 52) | fraction,
            i64::from(exponent_bits) - 1023 - 52,
        )
    };
    Some((negative, mantissa, exponent))
}

#[derive(Clone)]
struct SignedBig<const N: usize> {
    negative: bool,
    magnitude: Big<N>,
}

impl<const N: usize> SignedBig<N> {
    fn zero() -> Self {
        Self {
            negative: false,
            magnitude: Big::zero(),
        }
    }

    fn from_unsigned(negative: bool, magnitude: Big<N>) -> Self {
        Self {
            negative: negative && !magnitude.is_zero(),
            magnitude,
        }
    }

    fn from_positive(magnitude: &Big<N>) -> Self {
        Self::from_unsigned(false, magnitude.clone())
    }

    fn is_zero(&self) -> bool {
        self.magnitude.is_zero()
    }

    fn add_unsigned(&mut self, negative: bool, magnitude: &Big<N>) -> Result<(), MomentError> {
        let incoming = Self::from_unsigned(negative, magnitude.clone());
        self.add_signed(&incoming)
    }

    fn sub_unsigned(&mut self, magnitude: &Big<N>) -> Result<(), MomentError> {
        let incoming = Self::from_unsigned(true, magnitude.clone());
        self.add_signed(&incoming)
    }

    fn sub_signed(&mut self, other: &Self) -> Result<(), MomentError> {
        let incoming = Self::from_unsigned(!other.negative, other.magnitude.clone());
        self.add_signed(&incoming)
    }

    fn add_signed(&mut self, other: &Self) -> Result<(), MomentError> {
        if other.is_zero() {
            return Ok(());
        }
        if self.is_zero() {
            self.negative = other.negative;
            self.magnitude = other.magnitude.clone();
            return Ok(());
        }
        if self.negative == other.negative {
            self.magnitude.add_assign(&other.magnitude)?;
            return Ok(());
        }
        match self.magnitude.cmp(&other.magnitude) {
            Ordering::Greater => self.magnitude.sub_assign(&other.magnitude),
            Ordering::Equal => {
                self.magnitude.clear();
                self.negative = false;
            },
            Ordering::Less => {
                let mut difference = other.magnitude.clone();
                difference.sub_assign(&self.magnitude);
                self.magnitude = difference;
                self.negative = other.negative;
            },
        }
        Ok(())
    }

    fn mul_small(&mut self, factor: u64) -> Result<(), MomentError> {
        self.magnitude.mul_small(factor)
    }
}

fn to_calc_signed<const N: usize>(value: &SignedDyadic<N>) -> SignedBig<CALC_LIMBS> {
    let mut magnitude = Big::<CALC_LIMBS>::zero();
    for index in 0..value.magnitude.len {
        magnitude.limbs[index] = value.magnitude.limbs[index];
    }
    magnitude.len = value.magnitude.len;
    SignedBig::from_unsigned(value.negative, magnitude)
}

#[derive(Clone)]
struct Big<const N: usize> {
    limbs: [u64; N],
    len: usize,
}

impl<const N: usize> Big<N> {
    fn zero() -> Self {
        Self {
            limbs: [0; N],
            len: 0,
        }
    }

    fn from_small(value: u64) -> Self {
        if value == 0 {
            return Self::zero();
        }
        let mut result = Self::zero();
        result.limbs[0] = value;
        result.len = 1;
        result
    }

    fn is_zero(&self) -> bool {
        self.len == 0
    }

    fn clear(&mut self) {
        self.len = 0;
    }

    fn bit_length(&self) -> usize {
        self.len
            .checked_mul(64)
            .unwrap_or(usize::MAX)
            .saturating_sub(
                self.len
                    .checked_sub(1)
                    .and_then(|index| self.limbs.get(index))
                    .map_or(64, |limb| limb.leading_zeros() as usize),
            )
    }

    fn bit(&self, index: usize) -> bool {
        let word = index / 64;
        word < self.len && ((self.limbs[word] >> (index % 64)) & 1) != 0
    }

    fn cmp(&self, other: &Self) -> Ordering {
        if self.len != other.len {
            return self.len.cmp(&other.len);
        }
        for index in (0..self.len).rev() {
            if self.limbs[index] != other.limbs[index] {
                return self.limbs[index].cmp(&other.limbs[index]);
            }
        }
        Ordering::Equal
    }

    fn normalize(&mut self) {
        while self.len != 0 && self.limbs[self.len - 1] == 0 {
            self.len -= 1;
        }
    }

    fn mul_small(&mut self, factor: u64) -> Result<(), MomentError> {
        if factor == 0 || self.is_zero() {
            if factor == 0 {
                self.clear();
            }
            return Ok(());
        }
        if factor == 1 {
            return Ok(());
        }
        let mut carry = 0_u128;
        for index in 0..self.len {
            let product = (self.limbs[index] as u128) * (factor as u128) + carry;
            self.limbs[index] = product as u64;
            carry = product >> 64;
        }
        if carry != 0 {
            if self.len == N {
                return Err(MomentError::Number);
            }
            self.limbs[self.len] = carry as u64;
            self.len += 1;
        }
        Ok(())
    }

    fn add_assign(&mut self, other: &Self) -> Result<(), MomentError> {
        let length = self.len.max(other.len);
        if length > N {
            return Err(MomentError::Number);
        }
        let mut carry = 0_u64;
        for index in 0..length {
            let first = if index < self.len {
                self.limbs[index]
            } else {
                0
            };
            let second = if index < other.len {
                other.limbs[index]
            } else {
                0
            };
            let (sum, overflow_first) = first.overflowing_add(second);
            let (sum, overflow_second) = sum.overflowing_add(carry);
            self.limbs[index] = sum;
            carry = u64::from(overflow_first || overflow_second);
        }
        self.len = length;
        if carry != 0 {
            if self.len == N {
                return Err(MomentError::Number);
            }
            self.limbs[self.len] = carry;
            self.len += 1;
        }
        Ok(())
    }

    fn sub_assign(&mut self, other: &Self) {
        debug_assert!(self.cmp(other) != Ordering::Less);
        let mut borrow = 0_u64;
        for index in 0..self.len {
            let first = self.limbs[index];
            let second = if index < other.len {
                other.limbs[index]
            } else {
                0
            };
            let (difference, borrow_first) = first.overflowing_sub(second);
            let (difference, borrow_second) = difference.overflowing_sub(borrow);
            self.limbs[index] = difference;
            borrow = u64::from(borrow_first || borrow_second);
        }
        self.normalize();
    }

    fn assign_shifted(&mut self, other: &Self, bits: usize) -> Result<(), MomentError> {
        if other.is_zero() {
            self.clear();
            return Ok(());
        }
        let word_shift = bits / 64;
        let bit_shift = bits % 64;
        let extra = usize::from(bit_shift != 0);
        let new_len = other
            .len
            .checked_add(word_shift)
            .and_then(|length| length.checked_add(extra))
            .ok_or(MomentError::Number)?;
        if new_len > N {
            return Err(MomentError::Number);
        }
        for limb in &mut self.limbs[..new_len] {
            *limb = 0;
        }
        if bit_shift == 0 {
            for index in 0..other.len {
                self.limbs[index + word_shift] = other.limbs[index];
            }
        } else {
            let mut carry = 0_u64;
            for index in 0..other.len {
                let value = other.limbs[index];
                self.limbs[index + word_shift] = (value << bit_shift) | carry;
                carry = value >> (64 - bit_shift);
            }
            self.limbs[other.len + word_shift] = carry;
        }
        self.len = new_len;
        self.normalize();
        Ok(())
    }

    fn add_shifted_assign(&mut self, other: &Self, bits: usize) -> Result<(), MomentError> {
        if other.is_zero() {
            return Ok(());
        }
        let word_shift = bits / 64;
        let bit_shift = bits % 64;
        let mut carry = 0_u64;
        for index in 0..other.len {
            let value = other.limbs[index];
            let chunk = if bit_shift == 0 {
                value
            } else {
                let chunk = (value << bit_shift) | carry;
                carry = value >> (64 - bit_shift);
                chunk
            };
            self.add_limb_at(index + word_shift, chunk)?;
        }
        if bit_shift != 0 && carry != 0 {
            self.add_limb_at(other.len + word_shift, carry)?;
        }
        Ok(())
    }

    fn sub_shifted_assign(&mut self, other: &Self, bits: usize) -> Result<(), MomentError> {
        debug_assert!(cmp_shifted(self, other, bits) != Ordering::Less);
        if other.is_zero() {
            return Ok(());
        }
        let word_shift = bits / 64;
        let bit_shift = bits % 64;
        let mut borrow = 0_u64;
        let mut carry = 0_u64;
        for index in 0..other.len {
            let value = other.limbs[index];
            let chunk = if bit_shift == 0 {
                value
            } else {
                let chunk = (value << bit_shift) | carry;
                carry = value >> (64 - bit_shift);
                chunk
            };
            self.sub_limb_at(index + word_shift, chunk, &mut borrow);
        }
        let mut next = other.len + word_shift;
        if bit_shift != 0 && carry != 0 {
            self.sub_limb_at(next, carry, &mut borrow);
            next += 1;
        }
        // The shifted term may end below a nonzero high limb of `self`.  A
        // borrow from its highest occupied limb must continue through those
        // limbs; stopping at the term width silently corrupts a wide signed
        // cancellation in release builds.
        while borrow != 0 && next < self.len {
            self.sub_limb_at(next, 0, &mut borrow);
            next += 1;
        }
        if borrow != 0 {
            return Err(MomentError::Number);
        }
        self.normalize();
        Ok(())
    }

    fn add_limb_at(&mut self, index: usize, value: u64) -> Result<(), MomentError> {
        if value == 0 {
            return Ok(());
        }
        if index >= N {
            return Err(MomentError::Number);
        }
        while self.len <= index {
            self.limbs[self.len] = 0;
            self.len += 1;
        }
        let mut carry = value;
        let mut position = index;
        while carry != 0 {
            if position >= N {
                return Err(MomentError::Number);
            }
            if position >= self.len {
                self.limbs[position] = 0;
                self.len = position + 1;
            }
            let (sum, overflow) = self.limbs[position].overflowing_add(carry);
            self.limbs[position] = sum;
            carry = u64::from(overflow);
            position += 1;
        }
        Ok(())
    }

    fn sub_limb_at(&mut self, index: usize, value: u64, borrow: &mut u64) {
        debug_assert!(index < self.len);
        let (difference, borrow_first) = self.limbs[index].overflowing_sub(value);
        let (difference, borrow_second) = difference.overflowing_sub(*borrow);
        self.limbs[index] = difference;
        *borrow = u64::from(borrow_first || borrow_second);
    }

    fn shift_left(&mut self, bits: usize) -> Result<(), MomentError> {
        if bits == 0 || self.is_zero() {
            return Ok(());
        }
        let word_shift = bits / 64;
        let bit_shift = bits % 64;
        let old_len = self.len;
        let extra = usize::from(bit_shift != 0);
        let new_len = old_len
            .checked_add(word_shift)
            .and_then(|length| length.checked_add(extra))
            .ok_or(MomentError::Number)?;
        if new_len > N {
            return Err(MomentError::Number);
        }
        if word_shift != 0 {
            for index in (0..old_len).rev() {
                self.limbs[index + word_shift] = self.limbs[index];
            }
            for index in 0..word_shift {
                self.limbs[index] = 0;
            }
        }
        if bit_shift != 0 {
            let mut carry = 0_u64;
            for index in word_shift..(old_len + word_shift) {
                let value = self.limbs[index];
                self.limbs[index] = (value << bit_shift) | carry;
                carry = value >> (64 - bit_shift);
            }
            self.limbs[old_len + word_shift] = carry;
        }
        self.len = new_len;
        self.normalize();
        Ok(())
    }

    fn mul(&self, other: &Self) -> Result<Self, MomentError> {
        if self.is_zero() || other.is_zero() {
            return Ok(Self::zero());
        }
        let mut result = Self::zero();
        for i in 0..self.len {
            let mut carry = 0_u128;
            for j in 0..other.len {
                let index = i + j;
                if index >= N {
                    return Err(MomentError::Number);
                }
                let product = (self.limbs[i] as u128) * (other.limbs[j] as u128)
                    + (result.limbs[index] as u128)
                    + carry;
                result.limbs[index] = product as u64;
                carry = product >> 64;
            }
            let mut index = i + other.len;
            while carry != 0 {
                if index >= N {
                    return Err(MomentError::Number);
                }
                let (sum, overflow) = result.limbs[index].overflowing_add(carry as u64);
                result.limbs[index] = sum;
                carry = (carry >> 64) + u128::from(overflow);
                index += 1;
            }
        }
        result.len = (self.len + other.len).min(N);
        result.normalize();
        Ok(result)
    }

    fn normalized(&self) -> (f64, i64) {
        debug_assert!(!self.is_zero());
        let bits = self.bit_length();
        let take = bits.min(53);
        let shift = bits - take;
        let mut top = 0_u64;
        for offset in 0..take {
            if self.bit(shift + offset) {
                top |= 1_u64 << offset;
            }
        }
        let mantissa = top as f64 / 2_f64.powi((take - 1) as i32);
        (mantissa, bits as i64 - 1)
    }
}

fn multiply<const N: usize>(left: &Big<N>, right: &Big<N>) -> Result<Big<N>, MomentError> {
    left.mul(right)
}

fn cmp_shifted<const N: usize>(left: &Big<N>, right: &Big<N>, shift: usize) -> Ordering {
    let left_bits = left.bit_length();
    let right_bits = right.bit_length().saturating_add(shift);
    if left_bits != right_bits {
        return left_bits.cmp(&right_bits);
    }
    for index in (0..left_bits).rev() {
        let left_bit = left.bit(index);
        let right_bit = index >= shift && right.bit(index - shift);
        if left_bit != right_bit {
            return left_bit.cmp(&right_bit);
        }
    }
    Ordering::Equal
}

fn round_signed_ratio(
    numerator: &SignedBig<CALC_LIMBS>,
    denominator: &Big<CALC_LIMBS>,
    base: i64,
) -> Result<f64, MomentError> {
    if numerator.is_zero() {
        return Ok(0.0);
    }
    round_ratio(&numerator.magnitude, denominator, base, numerator.negative)
}

fn round_ratio(
    numerator: &Big<CALC_LIMBS>,
    denominator: &Big<CALC_LIMBS>,
    base: i64,
    negative: bool,
) -> Result<f64, MomentError> {
    if numerator.is_zero() {
        return Ok(0.0);
    }
    if denominator.is_zero() {
        return Err(MomentError::Number);
    }
    let mut exponent = numerator.bit_length() as i64 - denominator.bit_length() as i64;
    if exponent >= 0 {
        let shift = usize::try_from(exponent).map_err(|_| MomentError::Number)?;
        if cmp_shifted(numerator, denominator, shift) == Ordering::Less {
            exponent -= 1;
        }
    } else {
        let shift = usize::try_from(-exponent).map_err(|_| MomentError::Number)?;
        if cmp_shifted(denominator, numerator, shift) == Ordering::Greater {
            exponent -= 1;
        }
    }

    let mut scaled_numerator = numerator.clone();
    let mut scaled_denominator = denominator.clone();
    if exponent >= 0 {
        scaled_denominator
            .shift_left(usize::try_from(exponent).map_err(|_| MomentError::Number)?)?;
    } else {
        scaled_numerator
            .shift_left(usize::try_from(-exponent).map_err(|_| MomentError::Number)?)?;
    }
    if scaled_numerator.cmp(&scaled_denominator) == Ordering::Less {
        return Err(MomentError::Number);
    }
    scaled_numerator.sub_assign(&scaled_denominator);

    let mut mantissa = 1_u64 << 52;
    for bit in (0..52).rev() {
        scaled_numerator.shift_left(1)?;
        if scaled_numerator.cmp(&scaled_denominator) != Ordering::Less {
            scaled_numerator.sub_assign(&scaled_denominator);
            mantissa |= 1_u64 << bit;
        }
    }
    scaled_numerator.shift_left(1)?;
    let guard = scaled_numerator.cmp(&scaled_denominator) != Ordering::Less;
    if guard {
        scaled_numerator.sub_assign(&scaled_denominator);
    }
    if guard && (!scaled_numerator.is_zero() || mantissa & 1 != 0) {
        mantissa += 1;
        if mantissa == 1_u64 << 53 {
            mantissa = 1_u64 << 52;
            exponent += 1;
        }
    }

    let unbiased = base.checked_add(exponent).ok_or(MomentError::Number)?;
    let sign = if negative { 1_u64 << 63 } else { 0 };
    if unbiased > 1023 {
        return Err(MomentError::Number);
    }
    if unbiased >= -1022 {
        let exponent_bits = u64::try_from(unbiased + 1023).map_err(|_| MomentError::Number)?;
        let fraction = mantissa - (1_u64 << 52);
        return Ok(f64::from_bits(sign | (exponent_bits << 52) | fraction));
    }

    // A subnormal has units of 2^-1074. Round the 53-bit normalized
    // significand after moving it down to that quantum.
    let shift = usize::try_from(-(unbiased + 1022)).map_err(|_| MomentError::Number)?;
    let mut subnormal = if shift >= 64 { 0 } else { mantissa >> shift };
    let guard = shift != 0 && shift - 1 < 64 && ((mantissa >> (shift - 1)) & 1) != 0;
    let sticky = shift > 1 && shift - 1 < 64 && (mantissa & ((1_u64 << (shift - 1)) - 1)) != 0;
    if guard && (sticky || subnormal & 1 != 0) {
        subnormal = subnormal.saturating_add(1);
    }
    if subnormal >= 1_u64 << 52 {
        return Ok(f64::from_bits(sign | (1_u64 << 52)));
    }
    Ok(f64::from_bits(sign | subnormal))
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

fn normalize_mantissa(value: &mut f64, exponent: &mut i64) -> Result<(), MomentError> {
    if *value == 0.0 {
        return Ok(());
    }
    if !value.is_finite() || *value < 0.0 {
        return Err(MomentError::Number);
    }
    while *value >= 2.0 {
        *value *= 0.5;
        *exponent = exponent.checked_add(1).ok_or(MomentError::Number)?;
    }
    while *value < 1.0 {
        *value *= 2.0;
        *exponent = exponent.checked_sub(1).ok_or(MomentError::Number)?;
    }
    Ok(())
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
    normalize_mantissa(&mut quotient, &mut exponent)?;
    Ok(NormalizedBinary {
        negative: third.negative,
        mantissa: quotient,
        exponent,
    })
}

fn scale_power_of_two(mut value: f64, mut exponent: i64) -> f64 {
    while exponent > 1023 {
        value *= f64::from_bits(0x7fe0_0000_0000_0000);
        if !value.is_finite() {
            return value;
        }
        exponent -= 1023;
    }
    while exponent < -1074 {
        value *= f64::from_bits(1);
        if value == 0.0 {
            return value;
        }
        exponent += 1074;
    }
    if exponent >= -1022 {
        value * f64::from_bits(((exponent + 1023) as u64) << 52)
    } else if exponent >= -1074 {
        value * f64::from_bits(1_u64 << (exponent + 1074))
    } else {
        0.0
    }
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
