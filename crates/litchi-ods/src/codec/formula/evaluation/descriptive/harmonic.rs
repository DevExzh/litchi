//! Bounded signed binary kernels for the OpenFormula `HARMEAN` reducer.
//!
//! `HARMEAN` admits finite, non-zero numbers and returns `n / sum(1 / x)`. A
//! direct reciprocal is not a safe intermediate: the reciprocal of the
//! smallest subnormal overflows while the reciprocal of a large finite value
//! underflows. The first pass therefore keeps rounded reciprocal significands
//! in a fixed binary quantum with a conservative outward error bound. When
//! that interval cannot prove a sign and an accurate finite result, callers
//! replay their existing scalar stack or reference source through
//! [`ExactHarmonicAccumulator`]. No input values are retained in this module.
//!
//! The exact replay represents
//!
//! ```text
//! sum(1 / x) = N / D * 2^base
//! ```
//!
//! where `D` is odd. For `x = sign * m * 2^e`, with odd `m`, a new
//! denominator is admitted by `g = gcd(D mod m, m)`, `f = m / g`, followed by
//! `N *= f`, `D *= f`, and the signed addend `(old D / g) * 2^(-e-base)`.
//! Its limb vectors grow only as the checked caller profile admits them. A
//! profile refusal or allocator failure is typed; an exact tail is never
//! silently discarded.

use core::{
    cmp::Ordering,
    mem::{self, size_of},
};
use std::collections::TryReserveError;

#[path = "reciprocal_sum.rs"]
mod reciprocal_sum;

/// Outcome of the bounded floating-point first pass.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) enum HarmonicFirstPass {
    /// The retained interval proved this finite result.
    Value(f64),
    /// The caller must replay admitted values through the exact accumulator.
    Replay,
}

/// A failure produced by a pure harmonic kernel.
///
/// Formula-level domain failures remain distinct from bounded-evaluation and
/// allocator failures. Descriptive callers map `Resource` to their
/// evaluator's `Resource::Memory` diagnostic and preserve `Allocation` as a
/// typed evaluator allocation failure.
#[derive(Debug)]
pub(crate) enum HarmonicFailure {
    /// The input contains zero, no values were admitted, or the reciprocal
    /// sum is exactly zero.
    DivisionByZero,
    /// An input or the finite result is outside the scalar numeric domain.
    Number,
    /// The exact replay cannot retain the required magnitude under the
    /// caller-provided limb budget.
    Resource,
    /// The allocator rejected an aggregate-limb growth request. The original
    /// error is retained so evaluator callers can publish a typed allocation
    /// failure instead of converting it into a formula error or a resource
    /// limit.
    Allocation {
        /// The allocator's original failure.
        source: TryReserveError,
    },
}

/// Fixed-state signed harmonic first pass shared by scalar and value profiles.
#[derive(Clone, Copy, Debug)]
pub(crate) struct HarmonicKernel {
    count: u64,
    fast: reciprocal_sum::ReciprocalSum,
}

impl HarmonicKernel {
    /// Create an empty signed harmonic first pass.
    pub(crate) const fn new() -> Self {
        Self {
            count: 0,
            fast: reciprocal_sum::ReciprocalSum::new(),
        }
    }

    /// Return the number of admitted numeric values.
    pub(crate) const fn count(&self) -> u64 {
        self.count
    }

    /// Fold one finite, non-zero number into the first pass.
    pub(crate) fn push_number(&mut self, value: f64) -> Result<(), HarmonicFailure> {
        if !value.is_finite() {
            return Err(HarmonicFailure::Number);
        }
        decompose(value).ok_or(HarmonicFailure::DivisionByZero)?;
        self.count = self.count.checked_add(1).ok_or(HarmonicFailure::Number)?;
        self.fast.push(value);
        Ok(())
    }

    /// Finish the first pass, requesting exact replay when its interval is
    /// inconclusive.
    pub(crate) fn finish(&self) -> Result<HarmonicFirstPass, HarmonicFailure> {
        if self.count == 0 {
            return Err(HarmonicFailure::DivisionByZero);
        }
        Ok(match self.fast.try_finish(self.count) {
            Some(value) => HarmonicFirstPass::Value(value),
            None => HarmonicFirstPass::Replay,
        })
    }
}

/// Adaptive exact replay accumulator.
///
/// `max_limbs` is supplied by the caller after it has checked the input-derived
/// work and storage profile. The vectors retain aggregate binary limbs only;
/// their size is independent of the number of input objects once the required
/// exponent/denominator precision is admitted.
#[derive(Debug)]
pub(crate) struct ExactHarmonicAccumulator {
    count: u64,
    state: ExactReciprocalSum,
}

/// Return a checked exact-limb recommendation for an admitted input count.
///
/// The 2,200-bit constant covers the full binary64 reciprocal exponent span
/// and final count carry. The `53 * count` term bounds odd denominator growth;
/// this is an arithmetic ceiling only. Callers reserve actual limb growth
/// and charge occupied precision before each operation.
pub(crate) fn recommended_limbs(value_count: u64) -> Result<usize, HarmonicFailure> {
    let count = usize::try_from(value_count).map_err(|_| HarmonicFailure::Resource)?;
    let bits = count
        .checked_mul(53)
        .and_then(|bits| bits.checked_add(2_200))
        .ok_or(HarmonicFailure::Resource)?;
    bits.checked_add(63)
        .map(|bits| bits / 64)
        .ok_or(HarmonicFailure::Resource)
}

impl ExactHarmonicAccumulator {
    /// Create an exact replay state with a caller-selected limb ceiling.
    pub(crate) fn new(max_limbs: usize) -> Result<Self, HarmonicFailure> {
        if max_limbs == 0 {
            return Err(HarmonicFailure::Resource);
        }
        Ok(Self {
            count: 0,
            state: ExactReciprocalSum::new(max_limbs),
        })
    }

    /// Return the number of admitted values.
    pub(crate) const fn count(&self) -> u64 {
        self.count
    }

    /// Return a conservative checked work charge for the next exact push.
    ///
    /// Callers should charge this value before invoking [`Self::push_number`]
    /// so local work limits and cancellation remain visible at the admission
    /// boundary, and before the storage preflight scans any limbs. The charge
    /// follows occupied precision plus the fixed binary64 exponent span. The
    /// count-derived precision ceiling is not charged per input value.
    pub(crate) fn push_work_units(&self) -> Result<u64, HarmonicFailure> {
        self.state.push_work_units()
    }

    /// Return a conservative checked work charge for exact final conversion.
    pub(crate) fn finish_work_units(&self) -> Result<u64, HarmonicFailure> {
        self.state.finish_work_units()
    }

    /// Return the absolute storage reservation required before the next push.
    ///
    /// The profile includes currently retained aggregate capacities and the
    /// old-plus-new allocation peak for every vector that this value can grow.
    /// Callers reserve this amount before [`Self::push_number`] and retain the
    /// reservation across later pushes; no count-derived maximum is charged
    /// up front.
    pub(crate) fn push_storage_bytes(&self, value: f64) -> Result<usize, HarmonicFailure> {
        if !value.is_finite() {
            return Err(HarmonicFailure::Number);
        }
        let (_, mantissa, exponent) = decompose(value).ok_or(HarmonicFailure::DivisionByZero)?;
        self.state.push_storage_bytes(mantissa, -exponent)
    }

    /// Return the absolute storage reservation required before finalization.
    ///
    /// The result-numerator scratch vector is included, including its
    /// transient old-plus-new growth when the final count multiplication
    /// needs another limb.
    pub(crate) fn finish_storage_bytes(&self) -> Result<usize, HarmonicFailure> {
        self.state.finish_storage_bytes(self.count)
    }

    /// Fold one finite, non-zero value during the exact replay.
    pub(crate) fn push_number(&mut self, value: f64) -> Result<(), HarmonicFailure> {
        if !value.is_finite() {
            return Err(HarmonicFailure::Number);
        }
        let (negative, mantissa, exponent) =
            decompose(value).ok_or(HarmonicFailure::DivisionByZero)?;
        self.count = self.count.checked_add(1).ok_or(HarmonicFailure::Number)?;
        self.state.push(negative, mantissa, -exponent)
    }

    /// Finish the exact replay with a finite rounded result.
    pub(crate) fn finish(&mut self) -> Result<f64, HarmonicFailure> {
        if self.count == 0 {
            return Err(HarmonicFailure::DivisionByZero);
        }
        self.state.finish(self.count)
    }
}

/// Decompose a finite non-zero binary64 as `sign * odd_mantissa * 2^e`.
fn decompose(value: f64) -> Option<(bool, u64, i64)> {
    let bits = value.to_bits();
    let negative = (bits >> 63) != 0;
    let exponent_bits = ((bits >> 52) & 0x7ff) as u16;
    let fraction = bits & ((1_u64 << 52) - 1);
    let (significand, exponent) = if exponent_bits == 0 {
        if fraction == 0 {
            return None;
        }
        (fraction, -1074_i64)
    } else {
        (
            (1_u64 << 52) | fraction,
            i64::from(exponent_bits) - 1023 - 52,
        )
    };
    let trailing = i64::from(significand.trailing_zeros());
    Some((negative, significand >> trailing, exponent + trailing))
}

/// Exact odd-denominator replay state.
#[derive(Debug)]
struct ExactReciprocalSum {
    initialized: bool,
    base: i64,
    numerator: SignedMagnitude,
    denominator: BigMagnitude,
    quotient: BigMagnitude,
    result_numerator: BigMagnitude,
}

impl ExactReciprocalSum {
    fn new(max_limbs: usize) -> Self {
        Self {
            initialized: false,
            base: 0,
            numerator: SignedMagnitude::zero(max_limbs),
            denominator: BigMagnitude::zero(max_limbs),
            quotient: BigMagnitude::zero(max_limbs),
            result_numerator: BigMagnitude::zero(max_limbs),
        }
    }

    fn push_work_units(&self) -> Result<u64, HarmonicFailure> {
        // A binary64 reciprocal can move the common base by at most 2,097
        // bits, or 33 limbs. Sixteen passes cover storage preflight modulus,
        // quotient copy/modulus/division, multiplication, shifts, signed
        // comparison/addition, normalization, and allocation initialization.
        // The fixed shift allowance also covers limbs introduced by rebasing.
        // The count-derived maximum remains an admission ceiling only;
        // charging it for every streamed value would make a short denominator
        // quadratic in the input count.
        let occupied = self.occupied_limbs()?;
        let possible_shift = u64::from(MAX_BINARY64_SHIFT_LIMBS);
        occupied
            .checked_mul(16)
            .and_then(|work| work.checked_add(possible_shift.checked_mul(16)?))
            .and_then(|work| work.checked_add(32))
            .ok_or(HarmonicFailure::Resource)
    }

    fn finish_work_units(&self) -> Result<u64, HarmonicFailure> {
        let occupied = self.occupied_limbs()?;
        occupied
            .checked_add(2)
            .and_then(|work| work.checked_mul(8))
            .and_then(|work| work.checked_add(16))
            .ok_or(HarmonicFailure::Resource)
    }

    fn occupied_limbs(&self) -> Result<u64, HarmonicFailure> {
        let total = self
            .numerator
            .magnitude
            .limbs
            .len()
            .checked_add(self.denominator.limbs.len())
            .and_then(|total| total.checked_add(self.quotient.limbs.len()))
            .and_then(|total| total.checked_add(self.result_numerator.limbs.len()))
            .ok_or(HarmonicFailure::Resource)?;
        u64::try_from(total).map_err(|_| HarmonicFailure::Resource)
    }

    fn push_storage_bytes(
        &self,
        mantissa: u64,
        reciprocal_exponent: i64,
    ) -> Result<usize, HarmonicFailure> {
        // The real push initializes D to one before taking its modulus. Do
        // the same for the first simulated value; otherwise the first odd
        // mantissa would be incorrectly treated as already cancelled.
        let denominator_mod = if self.initialized {
            self.denominator.mod_small(mantissa)
        } else if mantissa == 1 {
            0
        } else {
            1
        };
        let gcd = gcd_u64(denominator_mod, mantissa);
        let factor = mantissa / gcd;
        let mut simulation = StorageSimulation::from_state(self)?;
        simulation.push(reciprocal_exponent, gcd, factor)?;
        Ok(simulation.peak_bytes)
    }

    fn finish_storage_bytes(&self, count: u64) -> Result<usize, HarmonicFailure> {
        let mut simulation = StorageSimulation::from_state(self)?;
        simulation.ensure(RESULT_NUMERATOR, self.denominator.limbs.len())?;
        if !self.denominator.limbs.is_empty() {
            simulation.multiply_small(RESULT_NUMERATOR, count)?;
        }
        Ok(simulation.peak_bytes)
    }

    fn push(
        &mut self,
        negative: bool,
        mantissa: u64,
        reciprocal_exponent: i64,
    ) -> Result<(), HarmonicFailure> {
        if !self.initialized {
            self.initialized = true;
            self.base = reciprocal_exponent;
            self.denominator.set_one()?;
        }

        if reciprocal_exponent < self.base {
            let shift = usize::try_from(self.base - reciprocal_exponent)
                .map_err(|_| HarmonicFailure::Resource)?;
            self.numerator.shift_left(shift)?;
            self.base = reciprocal_exponent;
        }

        self.quotient.copy_from(&self.denominator)?;
        let gcd = gcd_u64(self.quotient.mod_small(mantissa), mantissa);
        let factor = mantissa / gcd;
        self.quotient.divide_small(gcd)?;
        self.numerator.multiply_small(factor)?;
        self.denominator.multiply_small(factor)?;

        let shift = usize::try_from(reciprocal_exponent - self.base)
            .map_err(|_| HarmonicFailure::Resource)?;
        self.quotient.shift_left(shift)?;
        self.numerator.add_signed(negative, &mut self.quotient)
    }

    fn finish(&mut self, count: u64) -> Result<f64, HarmonicFailure> {
        if !self.initialized || self.numerator.is_zero() {
            return Err(HarmonicFailure::DivisionByZero);
        }
        self.result_numerator.copy_from(&self.denominator)?;
        self.result_numerator.multiply_small(count)?;
        let (numerator_mantissa, numerator_exponent) = self.result_numerator.normalized();
        let (denominator_mantissa, denominator_exponent) = self.numerator.magnitude.normalized();
        let mut quotient = numerator_mantissa / denominator_mantissa;
        let mut exponent = numerator_exponent - denominator_exponent - self.base;
        normalize_f64(&mut quotient, &mut exponent);
        let magnitude = scale_binary(quotient, exponent);
        if !magnitude.is_finite() {
            return Err(HarmonicFailure::Number);
        }
        Ok(if self.numerator.negative {
            -magnitude
        } else {
            magnitude
        })
    }
}

const NUMERATOR: usize = 0;
const DENOMINATOR: usize = 1;
const QUOTIENT: usize = 2;
const RESULT_NUMERATOR: usize = 3;
const MAX_BINARY64_SHIFT_LIMBS: u8 = 33;

/// Allocation-free simulation of one exact aggregate operation.
///
/// `Vec::try_reserve_exact` retains the old allocation until the replacement
/// succeeds. This state models logical lengths and capacities before a push,
/// computes each old-plus-new peak in operation order, and is discarded before
/// the real mutation starts. It therefore gives the evaluator an absolute
/// reservation target without allocating a second input-sized scratch state.
#[derive(Clone, Copy, Debug)]
struct StorageSimulation {
    initialized: bool,
    base: i64,
    lengths: [usize; 4],
    capacities: [usize; 4],
    max_limbs: usize,
    current_bytes: usize,
    peak_bytes: usize,
}

impl StorageSimulation {
    fn from_state(state: &ExactReciprocalSum) -> Result<Self, HarmonicFailure> {
        let lengths = [
            state.numerator.magnitude.limbs.len(),
            state.denominator.limbs.len(),
            state.quotient.limbs.len(),
            state.result_numerator.limbs.len(),
        ];
        let capacities = [
            state.numerator.magnitude.limbs.capacity(),
            state.denominator.limbs.capacity(),
            state.quotient.limbs.capacity(),
            state.result_numerator.limbs.capacity(),
        ];
        let mut current_bytes = 0usize;
        for capacity in capacities {
            current_bytes = current_bytes
                .checked_add(
                    capacity
                        .checked_mul(size_of::<u64>())
                        .ok_or(HarmonicFailure::Resource)?,
                )
                .ok_or(HarmonicFailure::Resource)?;
        }
        Ok(Self {
            initialized: state.initialized,
            base: state.base,
            lengths,
            capacities,
            max_limbs: state.denominator.max_limbs,
            current_bytes,
            peak_bytes: current_bytes,
        })
    }

    fn push(
        &mut self,
        reciprocal_exponent: i64,
        divisor: u64,
        factor: u64,
    ) -> Result<(), HarmonicFailure> {
        if !self.initialized {
            self.initialized = true;
            self.base = reciprocal_exponent;
            self.ensure(DENOMINATOR, 1)?;
        }

        if reciprocal_exponent < self.base {
            let shift = usize::try_from(
                self.base
                    .checked_sub(reciprocal_exponent)
                    .ok_or(HarmonicFailure::Resource)?,
            )
            .map_err(|_| HarmonicFailure::Resource)?;
            self.shift_left(NUMERATOR, shift)?;
            self.base = reciprocal_exponent;
        }

        self.ensure(QUOTIENT, self.lengths[DENOMINATOR])?;
        self.lengths[QUOTIENT] = self.lengths[DENOMINATOR];
        // Division can only shorten the copied quotient. Keeping its old
        // logical length is a conservative upper bound for all subsequent
        // shifts and numerator addition.
        self.divide_small(divisor)?;
        self.multiply_small(NUMERATOR, factor)?;
        self.multiply_small(DENOMINATOR, factor)?;

        let shift = usize::try_from(
            reciprocal_exponent
                .checked_sub(self.base)
                .ok_or(HarmonicFailure::Resource)?,
        )
        .map_err(|_| HarmonicFailure::Resource)?;
        self.shift_left(QUOTIENT, shift)?;
        if self.lengths[QUOTIENT] != 0 {
            let required = self.lengths[NUMERATOR]
                .max(self.lengths[QUOTIENT])
                .checked_add(1)
                .ok_or(HarmonicFailure::Resource)?;
            self.ensure(NUMERATOR, required)?;
        }
        Ok(())
    }

    fn ensure(&mut self, index: usize, length: usize) -> Result<(), HarmonicFailure> {
        if length > self.max_limbs {
            return Err(HarmonicFailure::Resource);
        }
        if length <= self.lengths[index] {
            return Ok(());
        }
        if length > self.capacities[index] {
            let new_bytes = length
                .checked_mul(size_of::<u64>())
                .ok_or(HarmonicFailure::Resource)?;
            let transient = self
                .current_bytes
                .checked_add(new_bytes)
                .ok_or(HarmonicFailure::Resource)?;
            self.peak_bytes = self.peak_bytes.max(transient);
            let old_bytes = self.capacities[index]
                .checked_mul(size_of::<u64>())
                .ok_or(HarmonicFailure::Resource)?;
            self.current_bytes = self
                .current_bytes
                .checked_sub(old_bytes)
                .and_then(|bytes| bytes.checked_add(new_bytes))
                .ok_or(HarmonicFailure::Resource)?;
            self.capacities[index] = length;
        }
        self.lengths[index] = length;
        Ok(())
    }

    fn multiply_small(&mut self, index: usize, factor: u64) -> Result<(), HarmonicFailure> {
        if factor <= 1 || self.lengths[index] == 0 {
            return Ok(());
        }
        let length = self.lengths[index]
            .checked_add(1)
            .ok_or(HarmonicFailure::Resource)?;
        self.ensure(index, length)
    }

    fn divide_small(&mut self, divisor: u64) -> Result<(), HarmonicFailure> {
        if divisor == 0 {
            return Err(HarmonicFailure::Number);
        }
        Ok(())
    }

    fn shift_left(&mut self, index: usize, bits: usize) -> Result<(), HarmonicFailure> {
        if bits == 0 || self.lengths[index] == 0 {
            return Ok(());
        }
        let word_shift = bits / 64;
        let extra = usize::from(bits % 64 != 0);
        let length = self.lengths[index]
            .checked_add(word_shift)
            .and_then(|length| length.checked_add(extra))
            .ok_or(HarmonicFailure::Resource)?;
        self.ensure(index, length)
    }
}

/// A positive adaptive magnitude stored least-significant limb first.
#[derive(Debug)]
struct BigMagnitude {
    limbs: Vec<u64>,
    max_limbs: usize,
}

impl BigMagnitude {
    fn zero(max_limbs: usize) -> Self {
        Self {
            limbs: Vec::new(),
            max_limbs,
        }
    }

    fn is_zero(&self) -> bool {
        self.limbs.is_empty()
    }

    fn set_one(&mut self) -> Result<(), HarmonicFailure> {
        self.ensure_len(1)?;
        self.limbs[0] = 1;
        Ok(())
    }

    fn ensure_len(&mut self, length: usize) -> Result<(), HarmonicFailure> {
        if length > self.max_limbs {
            return Err(HarmonicFailure::Resource);
        }
        if length > self.limbs.len() {
            let additional = length - self.limbs.len();
            self.limbs
                .try_reserve_exact(additional)
                .map_err(|source| HarmonicFailure::Allocation { source })?;
            self.limbs.resize(length, 0);
        }
        Ok(())
    }

    fn copy_from(&mut self, other: &Self) -> Result<(), HarmonicFailure> {
        self.ensure_len(other.limbs.len())?;
        self.limbs[..other.limbs.len()].copy_from_slice(&other.limbs);
        self.limbs.truncate(other.limbs.len());
        Ok(())
    }

    fn normalize(&mut self) {
        while self.limbs.last().copied() == Some(0) {
            self.limbs.pop();
        }
    }

    fn compare(&self, other: &Self) -> Ordering {
        if self.limbs.len() != other.limbs.len() {
            return self.limbs.len().cmp(&other.limbs.len());
        }
        for index in (0..self.limbs.len()).rev() {
            if self.limbs[index] != other.limbs[index] {
                return self.limbs[index].cmp(&other.limbs[index]);
            }
        }
        Ordering::Equal
    }

    fn multiply_small(&mut self, factor: u64) -> Result<(), HarmonicFailure> {
        if factor == 0 {
            self.limbs.clear();
            return Ok(());
        }
        if factor == 1 || self.is_zero() {
            return Ok(());
        }
        let mut carry = 0_u128;
        for limb in &mut self.limbs {
            let product = (*limb as u128) * (factor as u128) + carry;
            *limb = product as u64;
            carry = product >> 64;
        }
        if carry != 0 {
            let length = self
                .limbs
                .len()
                .checked_add(1)
                .ok_or(HarmonicFailure::Resource)?;
            self.ensure_len(length)?;
            self.limbs[length - 1] = carry as u64;
        }
        Ok(())
    }

    fn divide_small(&mut self, divisor: u64) -> Result<(), HarmonicFailure> {
        if divisor == 0 {
            return Err(HarmonicFailure::Number);
        }
        let mut remainder = 0_u128;
        for limb in self.limbs.iter_mut().rev() {
            let value = (remainder << 64) | *limb as u128;
            *limb = (value / divisor as u128) as u64;
            remainder = value % divisor as u128;
        }
        if remainder != 0 {
            return Err(HarmonicFailure::Number);
        }
        self.normalize();
        Ok(())
    }

    fn mod_small(&self, divisor: u64) -> u64 {
        if divisor <= 1 || self.is_zero() {
            return 0;
        }
        let mut remainder = 0_u128;
        for limb in self.limbs.iter().rev() {
            remainder = ((remainder << 64) | *limb as u128) % divisor as u128;
        }
        remainder as u64
    }

    fn shift_left(&mut self, bits: usize) -> Result<(), HarmonicFailure> {
        if bits == 0 || self.is_zero() {
            return Ok(());
        }
        let word_shift = bits / 64;
        let bit_shift = bits % 64;
        let old_len = self.limbs.len();
        let extra = usize::from(bit_shift != 0);
        let new_len = old_len
            .checked_add(word_shift)
            .and_then(|length| length.checked_add(extra))
            .ok_or(HarmonicFailure::Resource)?;
        self.ensure_len(new_len)?;
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
            for index in word_shift..(word_shift + old_len) {
                let limb = self.limbs[index];
                self.limbs[index] = (limb << bit_shift) | carry;
                carry = limb >> (64 - bit_shift);
            }
            self.limbs[old_len + word_shift] = carry;
            if carry == 0 {
                self.limbs.pop();
            }
        }
        Ok(())
    }

    fn add_assign(&mut self, other: &Self) -> Result<(), HarmonicFailure> {
        let old_len = self.limbs.len();
        let length = old_len.max(other.limbs.len());
        self.ensure_len(length)?;
        let mut carry = 0_u64;
        for index in 0..length {
            let first = self.limbs[index];
            let second = other.limbs.get(index).copied().unwrap_or(0);
            let (sum, overflow_first) = first.overflowing_add(second);
            let (sum, overflow_second) = sum.overflowing_add(carry);
            self.limbs[index] = sum;
            carry = u64::from(overflow_first || overflow_second);
        }
        if carry != 0 {
            self.ensure_len(length + 1)?;
            self.limbs[length] = carry;
        }
        Ok(())
    }

    fn subtract_assign(&mut self, other: &Self) {
        let mut borrow = 0_u64;
        for index in 0..self.limbs.len() {
            let first = self.limbs[index];
            let second = other.limbs.get(index).copied().unwrap_or(0);
            let (difference, borrow_first) = first.overflowing_sub(second);
            let (difference, borrow_second) = difference.overflowing_sub(borrow);
            self.limbs[index] = difference;
            borrow = u64::from(borrow_first || borrow_second);
        }
        self.normalize();
    }

    fn bit_length(&self) -> usize {
        self.limbs
            .last()
            .map(|limb| self.limbs.len() * 64 - limb.leading_zeros() as usize)
            .unwrap_or(0)
    }

    fn bit(&self, index: usize) -> bool {
        let word = index / 64;
        word < self.limbs.len() && (self.limbs[word] & (1_u64 << (index % 64))) != 0
    }

    fn top_bits(&self, count: usize) -> u64 {
        debug_assert!(count <= 53);
        let bits = self.bit_length();
        let shift = bits - count;
        let word = shift / 64;
        let offset = shift % 64;
        let mut value = self.limbs[word] >> offset;
        if offset != 0 && word + 1 < self.limbs.len() {
            value |= self.limbs[word + 1] << (64 - offset);
        }
        value & ((1_u64 << count) - 1)
    }

    fn any_below(&self, bits: usize) -> bool {
        if bits == 0 {
            return false;
        }
        let whole = bits / 64;
        if self.limbs[..whole.min(self.limbs.len())]
            .iter()
            .any(|limb| *limb != 0)
        {
            return true;
        }
        let remainder = bits % 64;
        remainder != 0
            && whole < self.limbs.len()
            && (self.limbs[whole] & ((1_u64 << remainder) - 1)) != 0
    }

    /// Return a rounded normalized mantissa in `[1, 2)` and its power of two.
    fn normalized(&self) -> (f64, i64) {
        debug_assert!(!self.is_zero());
        let bits = self.bit_length();
        let take = bits.min(53);
        let mut top = self.top_bits(take);
        let discarded = bits.saturating_sub(53);
        if discarded != 0 {
            let guard = self.bit(discarded - 1);
            let sticky = self.any_below(discarded - 1);
            if guard && (sticky || (top & 1) != 0) {
                top += 1;
                if top == (1_u64 << 53) {
                    return (1.0, bits as i64);
                }
            }
        }
        let mantissa = top as f64 / 2_f64.powi((take - 1) as i32);
        (mantissa, bits as i64 - 1)
    }
}

/// Signed magnitude wrapper used for exact cancellation.
#[derive(Debug)]
struct SignedMagnitude {
    negative: bool,
    magnitude: BigMagnitude,
}

impl SignedMagnitude {
    fn zero(max_limbs: usize) -> Self {
        Self {
            negative: false,
            magnitude: BigMagnitude::zero(max_limbs),
        }
    }

    fn is_zero(&self) -> bool {
        self.magnitude.is_zero()
    }

    fn multiply_small(&mut self, factor: u64) -> Result<(), HarmonicFailure> {
        self.magnitude.multiply_small(factor)
    }

    fn shift_left(&mut self, bits: usize) -> Result<(), HarmonicFailure> {
        self.magnitude.shift_left(bits)
    }

    fn add_signed(
        &mut self,
        negative: bool,
        value: &mut BigMagnitude,
    ) -> Result<(), HarmonicFailure> {
        if value.is_zero() {
            return Ok(());
        }
        if self.is_zero() {
            self.negative = negative;
            return self.magnitude.copy_from(value);
        }
        if self.negative == negative {
            self.magnitude.add_assign(value)
        } else {
            match self.magnitude.compare(value) {
                Ordering::Greater => self.magnitude.subtract_assign(value),
                Ordering::Equal => {
                    self.magnitude.limbs.clear();
                    self.negative = false;
                },
                Ordering::Less => {
                    value.subtract_assign(&self.magnitude);
                    mem::swap(&mut self.magnitude, value);
                    self.negative = negative;
                },
            }
            Ok(())
        }
    }
}

fn gcd_u64(mut first: u64, mut second: u64) -> u64 {
    while second != 0 {
        let remainder = first % second;
        first = second;
        second = remainder;
    }
    if first == 0 { 1 } else { first }
}

fn normalize_f64(value: &mut f64, exponent: &mut i64) {
    if *value == 0.0 || !value.is_finite() {
        return;
    }
    while *value >= 2.0 {
        *value *= 0.5;
        *exponent += 1;
    }
    while *value < 1.0 {
        *value *= 2.0;
        *exponent -= 1;
    }
}

fn scale_binary(mut value: f64, mut exponent: i64) -> f64 {
    while exponent > 1023 {
        value *= pow2(1023);
        exponent -= 1023;
        if !value.is_finite() {
            return value;
        }
    }
    while exponent < -1074 {
        value *= pow2(-1074);
        exponent += 1074;
        if value == 0.0 {
            return value;
        }
    }
    value * pow2(exponent)
}

fn pow2(exponent: i64) -> f64 {
    if exponent > 1023 {
        f64::INFINITY
    } else if exponent >= -1022 {
        f64::from_bits(((exponent + 1023) as u64) << 52)
    } else if exponent >= -1074 {
        f64::from_bits(1_u64 << (exponent + 1074))
    } else {
        0.0
    }
}
