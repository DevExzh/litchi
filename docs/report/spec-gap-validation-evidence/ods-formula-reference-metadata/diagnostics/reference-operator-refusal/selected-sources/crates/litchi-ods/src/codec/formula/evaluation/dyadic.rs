//! Fixed-width exact binary arithmetic shared by descriptive and paired reducers.
//!
//! Every admitted finite binary64 is decomposed into an integer significand and
//! an exact power of two.  The integer containers below have compile-time
//! widths; arithmetic returns a typed overflow instead of growing or dropping
//! limbs.  The evaluator owns resource charging around calls into this module.

use core::cmp::Ordering;

/// The least binary exponent used by an exact binary64 decomposition.
pub(super) const MIN_BINARY_EXPONENT: i64 = -1074;

/// An arithmetic result that cannot fit the selected fixed-width container.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum ArithmeticError {
    Overflow,
}

#[derive(Clone, Copy, Debug)]
pub(super) struct SignedDyadic<const N: usize> {
    pub(super) negative: bool,
    pub(super) magnitude: Big<N>,
}

impl<const N: usize> SignedDyadic<N> {
    pub(super) fn zero() -> Self {
        Self {
            negative: false,
            magnitude: Big::zero(),
        }
    }

    pub(super) fn add_power(
        &mut self,
        negative: bool,
        mantissa: u64,
        exponent: i64,
        power: usize,
    ) -> Result<(), ArithmeticError> {
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
            .ok_or(ArithmeticError::Overflow)?;
        let term_negative = power % 2 == 1 && negative;
        self.add_shifted_term(term_negative, &term, shift)
    }

    pub(super) fn add_shifted_term(
        &mut self,
        term_negative: bool,
        term: &Big<N>,
        shift: usize,
    ) -> Result<(), ArithmeticError> {
        if term.is_zero() {
            return Ok(());
        }

        if self.magnitude.is_zero() {
            self.magnitude.assign_shifted(term, shift)?;
            self.negative = term_negative;
            return Ok(());
        }
        if self.negative == term_negative {
            self.magnitude.add_shifted_assign(term, shift)?;
            return Ok(());
        }
        match cmp_shifted(&self.magnitude, term, shift) {
            Ordering::Greater => self.magnitude.sub_shifted_assign(term, shift)?,
            Ordering::Equal => {
                self.magnitude.clear();
                self.negative = false;
            },
            Ordering::Less => {
                let old = self.magnitude;
                self.magnitude.assign_shifted(term, shift)?;
                self.magnitude.sub_assign(&old);
                self.negative = term_negative;
            },
        }
        Ok(())
    }
}

pub(super) fn decompose(value: f64) -> Option<(bool, u64, i64)> {
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

#[derive(Clone, Copy, Debug)]
pub(super) struct SignedBig<const N: usize> {
    pub(super) negative: bool,
    pub(super) magnitude: Big<N>,
}

impl<const N: usize> SignedBig<N> {
    pub(super) fn zero() -> Self {
        Self {
            negative: false,
            magnitude: Big::zero(),
        }
    }

    pub(super) fn from_unsigned(negative: bool, magnitude: Big<N>) -> Self {
        Self {
            negative: negative && !magnitude.is_zero(),
            magnitude,
        }
    }

    pub(super) fn from_positive(magnitude: &Big<N>) -> Self {
        Self::from_unsigned(false, *magnitude)
    }

    pub(super) fn is_zero(&self) -> bool {
        self.magnitude.is_zero()
    }

    pub(super) fn add_unsigned(
        &mut self,
        negative: bool,
        magnitude: &Big<N>,
    ) -> Result<(), ArithmeticError> {
        let incoming = Self::from_unsigned(negative, *magnitude);
        self.add_signed(&incoming)
    }

    pub(super) fn sub_unsigned(&mut self, magnitude: &Big<N>) -> Result<(), ArithmeticError> {
        let incoming = Self::from_unsigned(true, *magnitude);
        self.add_signed(&incoming)
    }

    pub(super) fn sub_signed(&mut self, other: &Self) -> Result<(), ArithmeticError> {
        let incoming = Self::from_unsigned(!other.negative, other.magnitude);
        self.add_signed(&incoming)
    }

    pub(super) fn add_signed(&mut self, other: &Self) -> Result<(), ArithmeticError> {
        if other.is_zero() {
            return Ok(());
        }
        if self.is_zero() {
            self.negative = other.negative;
            self.magnitude = other.magnitude;
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
                let mut difference = other.magnitude;
                difference.sub_assign(&self.magnitude);
                self.magnitude = difference;
                self.negative = other.negative;
            },
        }
        Ok(())
    }

    pub(super) fn mul_small(&mut self, factor: u64) -> Result<(), ArithmeticError> {
        self.magnitude.mul_small(factor)
    }
}

#[derive(Clone, Copy, Debug)]
pub(super) struct Big<const N: usize> {
    pub(super) limbs: [u64; N],
    pub(super) len: usize,
}

impl<const N: usize> Big<N> {
    pub(super) fn zero() -> Self {
        Self {
            limbs: [0; N],
            len: 0,
        }
    }

    pub(super) fn from_small(value: u64) -> Self {
        if value == 0 {
            return Self::zero();
        }
        let mut result = Self::zero();
        result.limbs[0] = value;
        result.len = 1;
        result
    }

    pub(super) fn is_zero(&self) -> bool {
        self.len == 0
    }

    pub(super) fn clear(&mut self) {
        self.len = 0;
    }

    pub(super) fn bit_length(&self) -> usize {
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

    pub(super) fn bit(&self, index: usize) -> bool {
        let word = index / 64;
        word < self.len && ((self.limbs[word] >> (index % 64)) & 1) != 0
    }

    pub(super) fn cmp(&self, other: &Self) -> Ordering {
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

    pub(super) fn normalize(&mut self) {
        while self.len != 0 && self.limbs[self.len - 1] == 0 {
            self.len -= 1;
        }
    }

    pub(super) fn mul_small(&mut self, factor: u64) -> Result<(), ArithmeticError> {
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
                return Err(ArithmeticError::Overflow);
            }
            self.limbs[self.len] = carry as u64;
            self.len += 1;
        }
        Ok(())
    }

    pub(super) fn add_assign(&mut self, other: &Self) -> Result<(), ArithmeticError> {
        let length = self.len.max(other.len);
        if length > N {
            return Err(ArithmeticError::Overflow);
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
                return Err(ArithmeticError::Overflow);
            }
            self.limbs[self.len] = carry;
            self.len += 1;
        }
        Ok(())
    }

    pub(super) fn sub_assign(&mut self, other: &Self) {
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

    pub(super) fn assign_shifted(
        &mut self,
        other: &Self,
        bits: usize,
    ) -> Result<(), ArithmeticError> {
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
            .ok_or(ArithmeticError::Overflow)?;
        if new_len > N {
            return Err(ArithmeticError::Overflow);
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

    pub(super) fn add_shifted_assign(
        &mut self,
        other: &Self,
        bits: usize,
    ) -> Result<(), ArithmeticError> {
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

    pub(super) fn sub_shifted_assign(
        &mut self,
        other: &Self,
        bits: usize,
    ) -> Result<(), ArithmeticError> {
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
            return Err(ArithmeticError::Overflow);
        }
        self.normalize();
        Ok(())
    }

    pub(super) fn add_limb_at(&mut self, index: usize, value: u64) -> Result<(), ArithmeticError> {
        if value == 0 {
            return Ok(());
        }
        if index >= N {
            return Err(ArithmeticError::Overflow);
        }
        while self.len <= index {
            self.limbs[self.len] = 0;
            self.len += 1;
        }
        let mut carry = value;
        let mut position = index;
        while carry != 0 {
            if position >= N {
                return Err(ArithmeticError::Overflow);
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

    pub(super) fn shift_left(&mut self, bits: usize) -> Result<(), ArithmeticError> {
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
            .ok_or(ArithmeticError::Overflow)?;
        if new_len > N {
            return Err(ArithmeticError::Overflow);
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

    pub(super) fn mul(&self, other: &Self) -> Result<Self, ArithmeticError> {
        if self.is_zero() || other.is_zero() {
            return Ok(Self::zero());
        }
        let mut result = Self::zero();
        for i in 0..self.len {
            let mut carry = 0_u128;
            for j in 0..other.len {
                let index = i + j;
                if index >= N {
                    return Err(ArithmeticError::Overflow);
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
                    return Err(ArithmeticError::Overflow);
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

    pub(in crate::codec::formula::evaluation) fn normalized(&self) -> (f64, i64) {
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

pub(super) fn multiply<const N: usize>(
    left: &Big<N>,
    right: &Big<N>,
) -> Result<Big<N>, ArithmeticError> {
    left.mul(right)
}

fn cmp_shifted<const N: usize>(left: &Big<N>, right: &Big<N>, shift: usize) -> Ordering {
    let left_bits = left.bit_length();
    let right_bits = if right.is_zero() {
        0
    } else {
        right.bit_length().saturating_add(shift)
    };
    if left_bits != right_bits {
        return left_bits.cmp(&right_bits);
    }

    // Compare whole limbs from most to least significant.  The shifted view
    // is synthesized in place, so division retains its fixed-width storage
    // and avoids rescanning every bit for every quotient position.
    let word_shift = shift / 64;
    let bit_shift = shift % 64;
    for index in (0..left.len).rev() {
        let shifted = shifted_limb(right, index, word_shift, bit_shift);
        if left.limbs[index] != shifted {
            return left.limbs[index].cmp(&shifted);
        }
    }
    Ordering::Equal
}

fn shifted_limb<const N: usize>(
    value: &Big<N>,
    index: usize,
    word_shift: usize,
    bit_shift: usize,
) -> u64 {
    if index < word_shift {
        return 0;
    }
    let source = index - word_shift;
    if bit_shift == 0 {
        return if source < value.len {
            value.limbs[source]
        } else {
            0
        };
    }

    let low = if source < value.len {
        value.limbs[source] << bit_shift
    } else {
        0
    };
    let high = if source != 0 && source - 1 < value.len {
        value.limbs[source - 1] >> (64 - bit_shift)
    } else {
        0
    };
    low | high
}

pub(super) fn round_signed_ratio<const N: usize>(
    numerator: &SignedBig<N>,
    denominator: &Big<N>,
    base: i64,
) -> Result<f64, ArithmeticError> {
    if numerator.is_zero() {
        return Ok(0.0);
    }
    round_ratio(&numerator.magnitude, denominator, base, numerator.negative)
}

/// Divide two fixed-width integers while retaining the complete remainder.
/// The quotient is built by restoring division, so this remains bounded by
/// the caller's compile-time width and never allocates a quotient vector.
fn div_rem<const N: usize>(
    numerator: &Big<N>,
    denominator: &Big<N>,
) -> Result<(Big<N>, Big<N>), ArithmeticError> {
    if denominator.is_zero() {
        return Err(ArithmeticError::Overflow);
    }
    let mut quotient = Big::zero();
    let mut remainder = *numerator;
    let numerator_bits = numerator.bit_length();
    let denominator_bits = denominator.bit_length();
    if numerator_bits < denominator_bits {
        return Ok((quotient, remainder));
    }
    let mut shift = numerator_bits - denominator_bits;
    loop {
        if cmp_shifted(&remainder, denominator, shift) != Ordering::Less {
            remainder.sub_shifted_assign(denominator, shift)?;
            quotient.add_limb_at(shift / 64, 1_u64 << (shift % 64))?;
        }
        if shift == 0 {
            break;
        }
        shift -= 1;
    }
    Ok((quotient, remainder))
}

fn rounded_integer<const N: usize>(
    numerator: &Big<N>,
    denominator: &Big<N>,
) -> Result<u64, ArithmeticError> {
    let (quotient, remainder) = div_rem(numerator, denominator)?;
    if quotient.bit_length() > 64 {
        return Err(ArithmeticError::Overflow);
    }
    let mut value = 0_u64;
    for index in 0..quotient.bit_length() {
        if quotient.bit(index) {
            value |= 1_u64 << index;
        }
    }

    // Compare 2*remainder with the denominator without shifting the
    // remainder.  This avoids an extra carry limb at the fixed-width edge.
    let mut complement = *denominator;
    complement.sub_assign(&remainder);
    let round_up = match remainder.cmp(&complement) {
        Ordering::Greater => true,
        Ordering::Equal => value & 1 != 0,
        Ordering::Less => false,
    };
    if round_up {
        value = value.checked_add(1).ok_or(ArithmeticError::Overflow)?;
    }
    Ok(value)
}

/// Round an exact signed dyadic ratio directly to binary64.  In particular,
/// the subnormal path divides at the `2^-1074` target quantum; it does not
/// round a normalized 53-bit significand and then round that result again.
pub(super) fn round_ratio<const N: usize>(
    numerator: &Big<N>,
    denominator: &Big<N>,
    base: i64,
    negative: bool,
) -> Result<f64, ArithmeticError> {
    if numerator.is_zero() {
        return Ok(0.0);
    }
    if denominator.is_zero() {
        return Err(ArithmeticError::Overflow);
    }
    let mut exponent = numerator.bit_length() as i64 - denominator.bit_length() as i64;
    if exponent >= 0 {
        let shift = usize::try_from(exponent).map_err(|_| ArithmeticError::Overflow)?;
        if cmp_shifted(numerator, denominator, shift) == Ordering::Less {
            exponent -= 1;
        }
    } else {
        let shift = usize::try_from(-exponent).map_err(|_| ArithmeticError::Overflow)?;
        if cmp_shifted(denominator, numerator, shift) == Ordering::Greater {
            exponent -= 1;
        }
    }

    let unbiased = base
        .checked_add(exponent)
        .ok_or(ArithmeticError::Overflow)?;
    let sign = if negative { 1_u64 << 63 } else { 0 };

    if unbiased < -1022 {
        // The target is an integer multiple of 2^-1074.  Preserve every
        // discarded bit in the quotient/remainder and apply ties-to-even at
        // this target quantum.
        let scale = base.checked_add(1074).ok_or(ArithmeticError::Overflow)?;
        let (scaled_numerator, scaled_denominator) = if scale >= 0 {
            let mut value = *numerator;
            value.shift_left(usize::try_from(scale).map_err(|_| ArithmeticError::Overflow)?)?;
            (value, *denominator)
        } else {
            let mut value = *denominator;
            value.shift_left(usize::try_from(-scale).map_err(|_| ArithmeticError::Overflow)?)?;
            (*numerator, value)
        };
        let subnormal = rounded_integer(&scaled_numerator, &scaled_denominator)?;
        if subnormal >= 1_u64 << 52 {
            return Ok(f64::from_bits(sign | (1_u64 << 52)));
        }
        return Ok(f64::from_bits(sign | subnormal));
    }

    let scale = 52_i64
        .checked_sub(exponent)
        .ok_or(ArithmeticError::Overflow)?;
    let (scaled_numerator, scaled_denominator) = if scale >= 0 {
        let mut value = *numerator;
        value.shift_left(usize::try_from(scale).map_err(|_| ArithmeticError::Overflow)?)?;
        (value, *denominator)
    } else {
        let mut value = *denominator;
        value.shift_left(usize::try_from(-scale).map_err(|_| ArithmeticError::Overflow)?)?;
        (*numerator, value)
    };
    let mut mantissa = rounded_integer(&scaled_numerator, &scaled_denominator)?;
    if !(1_u64 << 52..=1_u64 << 53).contains(&mantissa) {
        return Err(ArithmeticError::Overflow);
    }
    if mantissa == 1_u64 << 53 {
        mantissa >>= 1;
        exponent = exponent.checked_add(1).ok_or(ArithmeticError::Overflow)?;
    }
    let unbiased = base
        .checked_add(exponent)
        .ok_or(ArithmeticError::Overflow)?;
    if unbiased > 1023 {
        return Err(ArithmeticError::Overflow);
    }
    let exponent_bits = u64::try_from(unbiased + 1023).map_err(|_| ArithmeticError::Overflow)?;
    let fraction = mantissa - (1_u64 << 52);
    Ok(f64::from_bits(sign | (exponent_bits << 52) | fraction))
}

pub(super) fn normalize_mantissa(
    value: &mut f64,
    exponent: &mut i64,
) -> Result<(), ArithmeticError> {
    if *value == 0.0 {
        return Ok(());
    }
    if !value.is_finite() || *value < 0.0 {
        return Err(ArithmeticError::Overflow);
    }
    while *value >= 2.0 {
        *value *= 0.5;
        *exponent = exponent.checked_add(1).ok_or(ArithmeticError::Overflow)?;
    }
    while *value < 1.0 {
        *value *= 2.0;
        *exponent = exponent.checked_sub(1).ok_or(ArithmeticError::Overflow)?;
    }
    Ok(())
}

pub(super) fn scale_power_of_two(mut value: f64, mut exponent: i64) -> f64 {
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
    use super::{Big, cmp_shifted};
    use core::cmp::Ordering;

    fn from_limbs<const N: usize>(limbs: &[u64]) -> Big<N> {
        assert!(limbs.len() <= N);
        let mut value = Big::zero();
        value.limbs[..limbs.len()].copy_from_slice(limbs);
        value.len = limbs.len();
        value.normalize();
        value
    }

    fn bitwise_cmp<const N: usize>(left: &Big<N>, right: &Big<N>, shift: usize) -> Ordering {
        let left_bits = left.bit_length();
        let right_bits = if right.is_zero() {
            0
        } else {
            right.bit_length().saturating_add(shift)
        };
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

    #[test]
    fn shifted_comparison_matches_bitwise_boundaries() {
        let left = from_limbs::<8>(&[
            0x0123_4567_89ab_cdef,
            0xfedc_ba98_7654_3210,
            0x8000_0000_0000_0001,
        ]);
        let right = from_limbs::<8>(&[0x7654_3210_fedc_ba98, 0x0000_0000_0000_0001]);
        for shift in [0, 1, 2, 31, 32, 63, 64, 65, 127, 128, 129] {
            assert_eq!(
                cmp_shifted(&left, &right, shift),
                bitwise_cmp(&left, &right, shift),
                "shift {shift}"
            );
        }
    }

    #[test]
    fn shifted_comparison_handles_equal_cross_limb_values() {
        let right = from_limbs::<8>(&[
            0xffff_ffff_ffff_ffff,
            0x0123_4567_89ab_cdef,
            0x8000_0000_0000_0000,
        ]);
        for shift in [0, 1, 63, 64, 65, 127, 128] {
            let mut left = Big::zero();
            if right
                .len
                .checked_add(shift / 64)
                .and_then(|length| length.checked_add(usize::from(shift % 64 != 0)))
                .is_some_and(|length| length <= 8)
            {
                left.assign_shifted(&right, shift).unwrap();
                assert_eq!(cmp_shifted(&left, &right, shift), Ordering::Equal);
                assert_eq!(
                    cmp_shifted(&left, &right, shift),
                    bitwise_cmp(&left, &right, shift)
                );
            }
        }
    }

    #[test]
    fn shifted_comparison_handles_zero_and_near_capacity() {
        let zero = Big::<4>::zero();
        let one = from_limbs::<4>(&[1]);
        assert_eq!(cmp_shifted(&zero, &one, 0), Ordering::Less);
        assert_eq!(cmp_shifted(&zero, &one, 65), Ordering::Less);
        assert_eq!(cmp_shifted(&one, &zero, 65), Ordering::Greater);
        assert_eq!(cmp_shifted(&zero, &zero, 65), Ordering::Equal);

        let right = from_limbs::<4>(&[0, 0, 0, 1]);
        let left = from_limbs::<4>(&[u64::MAX, u64::MAX, u64::MAX, u64::MAX]);
        for shift in [0, 1, 63, 64, 65] {
            assert_eq!(
                cmp_shifted(&left, &right, shift),
                bitwise_cmp(&left, &right, shift),
                "shift {shift}"
            );
        }
    }
}
