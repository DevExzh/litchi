//! Fixed-width exact arithmetic for paired regression statistics.
//!
//! The value VM supplies an already paired stream of finite numbers.  This
//! kernel retains only raw sums and never stores the input sequence:
//! `Sx`, `Sy`, `Sxx`, `Syy`, and `Sxy`.  Centered numerators are built exactly
//! at finish time so large nearly-equal observations do not erase covariance
//! or regression residuals through a rounded mean.

use super::super::dyadic::{
    self, ArithmeticError, Big, MIN_BINARY_EXPONENT, SignedBig, SignedDyadic, decompose,
    normalize_mantissa, round_ratio, round_signed_ratio,
};

/// State widths follow the binary64 bounds documented beside `ExactPaired`.
const FIRST_STATE_LIMBS: usize = 34;
const SECOND_STATE_LIMBS: usize = 67;
// The centered products need 136 limbs.  Direct subnormal rounding may shift
// the denominator by 1,074 bits, so retain a further 17 limbs for that exact
// target-quantum division.
const CALC_LIMBS: usize = 160;

/// The raw sum quantum for one observation.
const FIRST_BASE: i64 = MIN_BINARY_EXPONENT;
/// The raw square and cross-product quantum.
const SECOND_BASE: i64 = MIN_BINARY_EXPONENT * 2;

/// Select only the raw sums needed by a paired reducer.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::codec::formula::evaluation) enum PairNeed {
    /// `COVAR`; retain `Sx`, `Sy`, and `Sxy`.
    Covariance,
    /// `SLOPE`, `INTERCEPT`, and `FORECAST`; add `Sxx`.
    Regression,
    /// `CORREL`, `PEARSON`, `RSQ`, and `STEYX`; add `Sxx` and `Syy`.
    Correlation,
}

impl PairNeed {
    fn needs_xx(self) -> bool {
        !matches!(self, Self::Covariance)
    }

    fn needs_yy(self) -> bool {
        matches!(self, Self::Correlation)
    }
}

/// Arithmetic/domain failures mapped by the value/scalar owner to formula
/// errors.  Provider, cancellation, and resource failures occur outside this
/// kernel and must never be converted into one of these variants.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::codec::formula::evaluation) enum PairError {
    Empty,
    MinimumCount,
    ZeroVariance,
    Number,
}

impl From<ArithmeticError> for PairError {
    fn from(_: ArithmeticError) -> Self {
        Self::Number
    }
}

/// Exact paired raw sums.  The largest possible first sum is 2,162 bits and
/// the largest second/cross sum is 4,260 bits when the checked count is a
/// `u64`; 34 and 67 limbs therefore cover every finite binary64 input and
/// count without an arbitrary precision fallback.  Centered values use 68
/// limbs and their products use at most 136; the 160-limb workspace also
/// covers the additional 1,074-bit subnormal-rounding shift.
#[derive(Debug)]
pub(in crate::codec::formula::evaluation) struct ExactPaired {
    need: PairNeed,
    count: u64,
    sx: SignedDyadic<FIRST_STATE_LIMBS>,
    sy: SignedDyadic<FIRST_STATE_LIMBS>,
    sxx: Big<SECOND_STATE_LIMBS>,
    syy: Big<SECOND_STATE_LIMBS>,
    sxy: SignedDyadic<SECOND_STATE_LIMBS>,
}

impl ExactPaired {
    /// Create an empty fixed-width paired state.
    pub(in crate::codec::formula::evaluation) fn with_need(need: PairNeed) -> Self {
        Self {
            need,
            count: 0,
            sx: SignedDyadic::zero(),
            sy: SignedDyadic::zero(),
            sxx: Big::zero(),
            syy: Big::zero(),
            sxy: SignedDyadic::zero(),
        }
    }

    /// Add one already-admitted pair exactly.  The value layer owns admission,
    /// pairing, and resolver error precedence; this method only rejects a
    /// non-finite numeric payload or a fixed-width arithmetic overflow.
    pub(in crate::codec::formula::evaluation) fn push_pair(
        &mut self,
        x: f64,
        y: f64,
    ) -> Result<(), PairError> {
        let (x_negative, x_mantissa, x_exponent) = decompose(x).ok_or(PairError::Number)?;
        let (y_negative, y_mantissa, y_exponent) = decompose(y).ok_or(PairError::Number)?;
        let next_count = self.count.checked_add(1).ok_or(PairError::Number)?;

        self.sx.add_power(x_negative, x_mantissa, x_exponent, 1)?;
        self.sy.add_power(y_negative, y_mantissa, y_exponent, 1)?;
        if self.need.needs_xx() {
            add_unsigned_power(&mut self.sxx, x_mantissa, x_exponent, 2)?;
        }
        if self.need.needs_yy() {
            add_unsigned_power(&mut self.syy, y_mantissa, y_exponent, 2)?;
        }
        add_cross_product(
            &mut self.sxy,
            x_negative ^ y_negative,
            x_mantissa,
            x_exponent,
            y_mantissa,
            y_exponent,
        )?;
        self.count = next_count;
        Ok(())
    }

    /// Fixed upper bound for one admitted pair.  The value bridge should
    /// charge this before calling [`Self::push_pair`]; it is deliberately
    /// below the scalar default step ceiling for ordinary 1,024-pair inputs.
    pub(in crate::codec::formula::evaluation) fn push_work_units(&self) -> u64 {
        let mut units = 8 + 2 * FIRST_STATE_LIMBS as u64 + SECOND_STATE_LIMBS as u64;
        if self.need.needs_xx() {
            units += SECOND_STATE_LIMBS as u64;
        }
        if self.need.needs_yy() {
            units += SECOND_STATE_LIMBS as u64;
        }
        units
    }

    /// Fixed finalization bound for the selected reducer.  This includes the
    /// centered products, exact ratio setup, and bounded quotient iterations.
    pub(in crate::codec::formula::evaluation) fn finish_work_units(&self) -> u64 {
        match self.need {
            PairNeed::Covariance => centered_work_units() + 2 * CALC_LIMBS as u64,
            PairNeed::Regression => regression_work_units(),
            PairNeed::Correlation => correlation_work_units(),
        }
    }

    /// Fixed preparation bound for an exact reusable regression fit.
    pub(in crate::codec::formula::evaluation) fn prepare_forecast_work_units(&self) -> u64 {
        regression_work_units()
    }

    /// Finish population covariance from one exact centered pass.
    pub(in crate::codec::formula::evaluation) fn finish_covar(&self) -> Result<f64, PairError> {
        let centered = self.centered(PairNeed::Covariance)?;
        let mut denominator = Big::<CALC_LIMBS>::from_small(self.count);
        denominator.mul_small(self.count)?;
        round_signed_ratio(&centered.cxy, &denominator, SECOND_BASE).map_err(Into::into)
    }

    /// Finish CORREL/PEARSON.  Their exact magnitude is the square root of
    /// `Cxy²/(Cxx*Cyy)`, which avoids first rounding a square-root denominator.
    pub(in crate::codec::formula::evaluation) fn finish_correl(&self) -> Result<f64, PairError> {
        let centered = self.centered(PairNeed::Correlation)?;
        ensure_variances(&centered)?;
        if centered.cxy.is_zero() {
            return Ok(0.0);
        }
        let cxx = centered.cxx.as_ref().ok_or(PairError::Number)?;
        let cyy = centered.cyy.as_ref().ok_or(PairError::Number)?;
        let denominator = dyadic::multiply(cxx, cyy)?;
        let numerator = dyadic::multiply(&centered.cxy.magnitude, &centered.cxy.magnitude)?;
        let value = sqrt_ratio(&numerator, &denominator, 0)?;
        Ok(if centered.cxy.negative { -value } else { value })
    }

    /// Finish RSQ from the exact squared correlation ratio.
    pub(in crate::codec::formula::evaluation) fn finish_rsq(&self) -> Result<f64, PairError> {
        let centered = self.centered(PairNeed::Correlation)?;
        ensure_variances(&centered)?;
        if centered.cxy.is_zero() {
            return Ok(0.0);
        }
        let numerator = dyadic::multiply(&centered.cxy.magnitude, &centered.cxy.magnitude)?;
        let cxx = centered.cxx.as_ref().ok_or(PairError::Number)?;
        let cyy = centered.cyy.as_ref().ok_or(PairError::Number)?;
        let denominator = dyadic::multiply(cxx, cyy)?;
        round_ratio(&numerator, &denominator, 0, false).map_err(Into::into)
    }

    /// Prepare the fixed exact fit used by SLOPE, INTERCEPT, and repeated
    /// FORECAST queries.  No input sequence is retained and centered sums are
    /// computed only once.
    pub(in crate::codec::formula::evaluation) fn prepare_forecast(
        &self,
    ) -> Result<ForecastFit, PairError> {
        let centered = self.centered(PairNeed::Regression)?;
        let cxx = centered.cxx.ok_or(PairError::Number)?;
        if cxx.is_zero() {
            return Err(PairError::ZeroVariance);
        }
        let sx = to_calc_signed(&self.sx);
        let sy = to_calc_signed(&self.sy);
        let cxy = centered.cxy;
        let left = multiply_signed_unsigned(&sy, &cxx)?;
        let right = multiply_signed(&sx, &cxy)?;
        let mut intercept_numerator = left;
        intercept_numerator.sub_signed(&right)?;
        let mut slope_numerator = cxy;
        slope_numerator.mul_small(self.count)?;
        let mut denominator = cxx;
        denominator.mul_small(self.count)?;
        Ok(ForecastFit {
            intercept_numerator,
            slope_numerator,
            denominator,
        })
    }

    pub(in crate::codec::formula::evaluation) fn finish_slope(&self) -> Result<f64, PairError> {
        self.prepare_forecast()?.finish_slope()
    }

    pub(in crate::codec::formula::evaluation) fn finish_intercept(&self) -> Result<f64, PairError> {
        self.prepare_forecast()?.finish_intercept()
    }

    pub(in crate::codec::formula::evaluation) fn finish_forecast(
        &self,
        x0: f64,
    ) -> Result<f64, PairError> {
        self.prepare_forecast()?.finish_forecast(x0)
    }

    /// Finish STEYX using the exact nonnegative residual determinant.
    pub(in crate::codec::formula::evaluation) fn finish_steyx(&self) -> Result<f64, PairError> {
        if self.count < 3 {
            return Err(PairError::MinimumCount);
        }
        let centered = self.centered(PairNeed::Correlation)?;
        let cxx = centered.cxx.as_ref().ok_or(PairError::Number)?;
        let cyy = centered.cyy.as_ref().ok_or(PairError::Number)?;
        // A constant dependent series has zero residuals, so STEYX is zero
        // even though Cyy is zero.  Only the independent variance is a
        // denominator and must be nonzero.
        if cxx.is_zero() {
            return Err(PairError::ZeroVariance);
        }
        let product = dyadic::multiply(cxx, cyy)?;
        let cross_square = dyadic::multiply(&centered.cxy.magnitude, &centered.cxy.magnitude)?;
        if product.cmp(&cross_square).is_lt() {
            return Err(PairError::Number);
        }
        let mut determinant = product;
        determinant.sub_assign(&cross_square);
        if determinant.is_zero() {
            return Ok(0.0);
        }
        let mut denominator = *cxx;
        denominator.mul_small(self.count)?;
        denominator.mul_small(self.count - 2)?;
        sqrt_ratio(&determinant, &denominator, SECOND_BASE)
    }

    fn centered(&self, need: PairNeed) -> Result<CenteredPair, PairError> {
        if self.count == 0 {
            return Err(PairError::Empty);
        }
        if need.needs_xx() && !self.need.needs_xx() {
            return Err(PairError::Number);
        }
        if need.needs_yy() && !self.need.needs_yy() {
            return Err(PairError::Number);
        }

        let sx = to_calc_signed(&self.sx);
        let sy = to_calc_signed(&self.sy);
        let sxy = to_calc_signed(&self.sxy);
        let cxx = if need.needs_xx() {
            Some(centered_square(&self.sxx, &sx, self.count)?)
        } else {
            None
        };
        let cyy = if need.needs_yy() {
            Some(centered_square(&self.syy, &sy, self.count)?)
        } else {
            None
        };

        let sx_sy = dyadic::multiply(&sx.magnitude, &sy.magnitude)?;
        let mut cxy = sxy;
        cxy.magnitude.mul_small(self.count)?;
        let sx_sy = SignedBig::from_unsigned(sx.negative ^ sy.negative, sx_sy);
        cxy.sub_signed(&sx_sy)?;

        Ok(CenteredPair { cxx, cyy, cxy })
    }
}

/// A reusable exact ordinary least-squares fit with an included constant.
/// The contract owner maps this profile to `INTERCEPT` and `FORECAST`.
#[derive(Clone, Copy, Debug)]
pub(in crate::codec::formula::evaluation) struct ForecastFit {
    intercept_numerator: SignedBig<CALC_LIMBS>,
    slope_numerator: SignedBig<CALC_LIMBS>,
    denominator: Big<CALC_LIMBS>,
}

impl ForecastFit {
    /// Fixed query bound for one scalar or projected FORECAST coordinate.
    pub(in crate::codec::formula::evaluation) const fn query_work_units() -> u64 {
        70 * 34 + 64
    }

    pub(in crate::codec::formula::evaluation) fn finish_slope(&self) -> Result<f64, PairError> {
        round_signed_ratio(&self.slope_numerator, &self.denominator, 0).map_err(Into::into)
    }

    pub(in crate::codec::formula::evaluation) fn finish_intercept(&self) -> Result<f64, PairError> {
        round_signed_ratio(&self.intercept_numerator, &self.denominator, FIRST_BASE)
            .map_err(Into::into)
    }

    pub(in crate::codec::formula::evaluation) fn finish_forecast(
        &self,
        x0: f64,
    ) -> Result<f64, PairError> {
        let (negative, mantissa, exponent) = decompose(x0).ok_or(PairError::Number)?;
        let mut x0_term = Big::<CALC_LIMBS>::from_small(mantissa);
        let shift = exponent
            .checked_sub(MIN_BINARY_EXPONENT)
            .and_then(|value| usize::try_from(value).ok())
            .ok_or(PairError::Number)?;
        x0_term.shift_left(shift)?;
        let x0_scaled = SignedBig::from_unsigned(negative, x0_term);
        let right = multiply_signed(&self.slope_numerator, &x0_scaled)?;
        let mut numerator = self.intercept_numerator;
        numerator.add_signed(&right)?;
        round_signed_ratio(&numerator, &self.denominator, FIRST_BASE).map_err(Into::into)
    }
}

const fn multiply_work_units(left: usize, right: usize) -> u64 {
    (left * right) as u64 + (left + right) as u64
}

const fn centered_work_units() -> u64 {
    3 * multiply_work_units(FIRST_STATE_LIMBS, FIRST_STATE_LIMBS) + 4 * CALC_LIMBS as u64
}

const fn regression_work_units() -> u64 {
    centered_work_units()
        + 2 * multiply_work_units(SECOND_STATE_LIMBS, FIRST_STATE_LIMBS)
        + 6 * CALC_LIMBS as u64
}

const fn correlation_work_units() -> u64 {
    centered_work_units()
        + 2 * multiply_work_units(SECOND_STATE_LIMBS, SECOND_STATE_LIMBS)
        + 8 * CALC_LIMBS as u64
}

struct CenteredPair {
    cxx: Option<Big<CALC_LIMBS>>,
    cyy: Option<Big<CALC_LIMBS>>,
    cxy: SignedBig<CALC_LIMBS>,
}

fn ensure_variances(centered: &CenteredPair) -> Result<(), PairError> {
    if centered.cxx.as_ref().is_none_or(Big::is_zero)
        || centered.cyy.as_ref().is_none_or(Big::is_zero)
    {
        return Err(PairError::ZeroVariance);
    }
    Ok(())
}

fn centered_square<const N: usize>(
    raw: &Big<N>,
    sum: &SignedBig<CALC_LIMBS>,
    count: u64,
) -> Result<Big<CALC_LIMBS>, PairError> {
    let sum_square = dyadic::multiply(&sum.magnitude, &sum.magnitude)?;
    let mut raw_term = to_calc_unsigned(raw);
    raw_term.mul_small(count)?;
    if raw_term.cmp(&sum_square).is_lt() {
        return Err(PairError::Number);
    }
    raw_term.sub_assign(&sum_square);
    Ok(raw_term)
}

fn to_calc_unsigned<const N: usize>(value: &Big<N>) -> Big<CALC_LIMBS> {
    let mut result = Big::<CALC_LIMBS>::zero();
    for index in 0..value.len {
        result.limbs[index] = value.limbs[index];
    }
    result.len = value.len;
    result
}

fn to_calc_signed<const N: usize>(value: &SignedDyadic<N>) -> SignedBig<CALC_LIMBS> {
    SignedBig::from_unsigned(value.negative, to_calc_unsigned(&value.magnitude))
}

fn multiply_signed_unsigned(
    left: &SignedBig<CALC_LIMBS>,
    right: &Big<CALC_LIMBS>,
) -> Result<SignedBig<CALC_LIMBS>, PairError> {
    let magnitude = dyadic::multiply(&left.magnitude, right)?;
    Ok(SignedBig::from_unsigned(left.negative, magnitude))
}

fn multiply_signed(
    left: &SignedBig<CALC_LIMBS>,
    right: &SignedBig<CALC_LIMBS>,
) -> Result<SignedBig<CALC_LIMBS>, PairError> {
    let magnitude = dyadic::multiply(&left.magnitude, &right.magnitude)?;
    Ok(SignedBig::from_unsigned(
        left.negative ^ right.negative,
        magnitude,
    ))
}

fn add_unsigned_power(
    sum: &mut Big<SECOND_STATE_LIMBS>,
    mantissa: u64,
    exponent: i64,
    power: usize,
) -> Result<(), PairError> {
    if mantissa == 0 {
        return Ok(());
    }
    let mut term = Big::<SECOND_STATE_LIMBS>::from_small(1);
    for _ in 0..power {
        term.mul_small(mantissa)?;
    }
    let shift = exponent
        .checked_sub(MIN_BINARY_EXPONENT)
        .and_then(|value| value.checked_mul(power as i64))
        .and_then(|value| usize::try_from(value).ok())
        .ok_or(PairError::Number)?;
    sum.add_shifted_assign(&term, shift)?;
    Ok(())
}

fn add_cross_product(
    sum: &mut SignedDyadic<SECOND_STATE_LIMBS>,
    negative: bool,
    x_mantissa: u64,
    x_exponent: i64,
    y_mantissa: u64,
    y_exponent: i64,
) -> Result<(), PairError> {
    if x_mantissa == 0 || y_mantissa == 0 {
        return Ok(());
    }
    let mut term = Big::<SECOND_STATE_LIMBS>::from_small(x_mantissa);
    term.mul_small(y_mantissa)?;
    let x_shift = x_exponent
        .checked_sub(MIN_BINARY_EXPONENT)
        .ok_or(PairError::Number)?;
    let y_shift = y_exponent
        .checked_sub(MIN_BINARY_EXPONENT)
        .ok_or(PairError::Number)?;
    let shift = x_shift
        .checked_add(y_shift)
        .and_then(|value| usize::try_from(value).ok())
        .ok_or(PairError::Number)?;
    sum.add_shifted_term(negative, &term, shift)?;
    Ok(())
}

/// Exponent-safe square root of an exact positive dyadic ratio.  Mantissas
/// are reduced only for the final libm square root; integer products and the
/// ratio exponent remain exact, so extreme magnitudes cannot overflow first.
fn sqrt_ratio(
    numerator: &Big<CALC_LIMBS>,
    denominator: &Big<CALC_LIMBS>,
    base: i64,
) -> Result<f64, PairError> {
    if numerator.is_zero() {
        return Ok(0.0);
    }
    if denominator.is_zero() {
        return Err(PairError::Number);
    }
    let (numerator_mantissa, numerator_exponent) = numerator.normalized();
    let (denominator_mantissa, denominator_exponent) = denominator.normalized();
    let mut quotient = numerator_mantissa / denominator_mantissa;
    let ratio_exponent = base
        .checked_add(numerator_exponent)
        .and_then(|value| value.checked_sub(denominator_exponent))
        .ok_or(PairError::Number)?;
    let mut exponent = ratio_exponent.div_euclid(2);
    if ratio_exponent.rem_euclid(2) != 0 {
        quotient *= 2.0;
    }
    quotient = quotient.sqrt();
    normalize_mantissa(&mut quotient, &mut exponent)?;
    let value = dyadic::scale_power_of_two(quotient, exponent);
    if !value.is_finite() {
        return Err(PairError::Number);
    }
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::{ExactPaired, ForecastFit, PairNeed};

    #[test]
    fn paired_centering_matches_basic_linear_data() {
        let mut state = ExactPaired::with_need(PairNeed::Correlation);
        assert!(state.push_work_units() > 0);
        for (x, y) in [(1.0, 2.0), (2.0, 4.0), (3.0, 6.0), (4.0, 8.0)] {
            state.push_pair(x, y).unwrap();
        }
        assert!(state.finish_work_units() > 0);
        assert_eq!(state.finish_covar().unwrap(), 2.5);
        assert_eq!(state.finish_correl().unwrap(), 1.0);
        assert_eq!(state.finish_rsq().unwrap(), 1.0);
        assert_eq!(state.finish_slope().unwrap(), 2.0);
        assert_eq!(state.finish_intercept().unwrap(), 0.0);
        assert_eq!(state.finish_forecast(5.0).unwrap(), 10.0);
        assert_eq!(state.finish_steyx().unwrap(), 0.0);
    }

    #[test]
    fn covariance_need_does_not_require_unused_square_sums() {
        let mut state = ExactPaired::with_need(PairNeed::Covariance);
        for (x, y) in [(1.0, 2.0), (2.0, 4.0), (3.0, 6.0), (4.0, 8.0)] {
            state.push_pair(x, y).unwrap();
        }
        assert_eq!(state.finish_covar().unwrap(), 2.5);
    }

    #[test]
    fn reusable_fit_keeps_exact_intercept_and_queries() {
        let mut state = ExactPaired::with_need(PairNeed::Regression);
        assert!(state.prepare_forecast_work_units() > 0);
        assert!(ForecastFit::query_work_units() > 0);
        for (x, y) in [(1.0, 5.0), (2.0, 3.0), (3.0, 1.0)] {
            state.push_pair(x, y).unwrap();
        }
        let fit = state.prepare_forecast().unwrap();
        assert_eq!(fit.finish_slope().unwrap(), -2.0);
        assert_eq!(fit.finish_intercept().unwrap(), 7.0);
        assert_eq!(fit.finish_forecast(4.0).unwrap(), -1.0);
        assert_eq!(fit.finish_forecast(-1.0).unwrap(), 9.0);
        assert_eq!(state.finish_slope().unwrap(), fit.finish_slope().unwrap());
        assert_eq!(
            state.finish_intercept().unwrap(),
            fit.finish_intercept().unwrap()
        );
    }

    #[test]
    fn paired_state_keeps_cancellation_exact() {
        let mut state = ExactPaired::with_need(PairNeed::Correlation);
        for (x, y) in [
            (4_503_599_627_370_496.0, 9_007_199_254_740_992.0),
            (4_503_599_627_370_497.0, 9_007_199_254_740_994.0),
        ] {
            state.push_pair(x, y).unwrap();
        }
        assert_eq!(state.finish_slope().unwrap(), 2.0);
        assert_eq!(state.finish_correl().unwrap(), 1.0);
    }

    #[test]
    fn steyx_constant_dependent_series_has_zero_residual() {
        let mut state = ExactPaired::with_need(PairNeed::Correlation);
        for x in [1.0, 2.0, 3.0, 4.0] {
            state.push_pair(x, 7.0).unwrap();
        }
        assert_eq!(state.finish_steyx().unwrap().to_bits(), 0.0_f64.to_bits());
    }
}
