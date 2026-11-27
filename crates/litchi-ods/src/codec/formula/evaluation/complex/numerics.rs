//! Overflow-aware numerical kernels for complex functions.
//!
//! The parent module owns the public representation and function dispatch.
//! These kernels deliberately do not construct an intermediate modulus or
//! quotient scale when that intermediate would overflow while the requested
//! result is still finite.

use super::{Complex, ScalarError};

const RATIO_TRIG_THRESHOLD: f64 = 20.0;

fn finite(real: f64, imaginary: f64, suffix: char) -> Result<Complex, ScalarError> {
    if real.is_finite() && imaginary.is_finite() && matches!(suffix, 'i' | 'j') {
        Ok(Complex {
            real,
            imaginary,
            suffix,
        })
    } else {
        Err(ScalarError::Number)
    }
}

/// Divide two finite complex values without forming either raw square sum.
pub(super) fn div(left: Complex, right: Complex) -> Result<Complex, ScalarError> {
    if right.real == 0.0 && right.imaginary == 0.0 {
        return Err(ScalarError::DivisionByZero);
    }
    if left.real == 0.0 && left.imaginary == 0.0 {
        return finite(0.0, 0.0, choose_suffix(left.suffix, right.suffix));
    }

    // Keep pure-axis denominators on direct component divisions. Normalizing
    // the whole numerator by its largest component would turn, for example,
    // 1e-308 / 1 into zero before the quotient is formed.
    if right.imaginary == 0.0 {
        return finite(
            left.real / right.real,
            left.imaginary / right.real,
            choose_suffix(left.suffix, right.suffix),
        );
    }
    if right.real == 0.0 {
        return finite(
            left.imaginary / right.imaginary,
            -(left.real / right.imaginary),
            choose_suffix(left.suffix, right.suffix),
        );
    }

    // Keep the numerator and denominator terms in binary form until the
    // final quotient. In particular, `d / c` can underflow even when
    // `a * d / (c^2 + d^2)` is a representable imaginary component.
    let denominator = combine_terms(
        product_term(right.real, right.real)?,
        product_term(right.imaginary, right.imaginary)?,
        false,
    )
    .ok_or(ScalarError::DivisionByZero)?;
    let real_numerator = combine_terms(
        product_term(left.real, right.real)?,
        product_term(left.imaginary, right.imaginary)?,
        false,
    );
    let imaginary_numerator = combine_terms(
        product_term(left.imaginary, right.real)?,
        product_term(left.real, right.imaginary)?,
        true,
    );
    let real = divide_term(real_numerator, denominator);
    let imaginary = divide_term(imaginary_numerator, denominator);
    finite(real, imaginary, choose_suffix(left.suffix, right.suffix))
}

/// Multiply finite complex values by combining the component products in
/// binary mantissa/exponent form. A direct `a * c - b * d` can overflow even
/// when the subtraction is finite, while normalizing all four components to a
/// common scale can erase a representable small component.
pub(super) fn mul(left: Complex, right: Complex) -> Result<Complex, ScalarError> {
    let real = sum_products(
        product_term(left.real, right.real)?,
        product_term(left.imaginary, right.imaginary)?,
        true,
    );
    let imaginary = sum_products(
        product_term(left.real, right.imaginary)?,
        product_term(left.imaginary, right.real)?,
        false,
    );
    finite(real, imaginary, choose_suffix(left.suffix, right.suffix))
}

/// Compute the complex exponential without materializing an overflowing
/// `exp(real)` before multiplying by the trigonometric coefficients.
pub(super) fn exp(value: Complex) -> Result<Complex, ScalarError> {
    let (sin_imaginary, cos_imaginary) = value.imaginary.sin_cos();
    finite(
        exponential_product(value.real, cos_imaginary),
        exponential_product(value.real, sin_imaginary),
        value.suffix,
    )
}

/// Compute the complex sine with separately scaled hyperbolic components.
pub(super) fn sin(value: Complex) -> Result<Complex, ScalarError> {
    let (sin_real, cos_real) = value.real.sin_cos();
    finite(
        hyperbolic_cosine_product(value.imaginary, sin_real),
        hyperbolic_sine_product(value.imaginary, cos_real),
        value.suffix,
    )
}

/// Compute the complex cosine with separately scaled hyperbolic components.
pub(super) fn cos(value: Complex) -> Result<Complex, ScalarError> {
    let (sin_real, cos_real) = value.real.sin_cos();
    finite(
        hyperbolic_cosine_product(value.imaginary, cos_real),
        -hyperbolic_sine_product(value.imaginary, sin_real),
        value.suffix,
    )
}

/// Compute the complex hyperbolic sine with separately scaled components.
pub(super) fn sinh(value: Complex) -> Result<Complex, ScalarError> {
    let (sin_imaginary, cos_imaginary) = value.imaginary.sin_cos();
    finite(
        hyperbolic_sine_product(value.real, cos_imaginary),
        hyperbolic_cosine_product(value.real, sin_imaginary),
        value.suffix,
    )
}

/// Compute the complex hyperbolic cosine with separately scaled components.
pub(super) fn cosh(value: Complex) -> Result<Complex, ScalarError> {
    let (sin_imaginary, cos_imaginary) = value.imaginary.sin_cos();
    finite(
        hyperbolic_cosine_product(value.real, cos_imaginary),
        hyperbolic_sine_product(value.real, sin_imaginary),
        value.suffix,
    )
}

/// Compute `ln(|real + imaginary*i|)` without requiring the modulus itself to
/// be representable as an `f64`.
pub(super) fn ln(value: Complex) -> Result<Complex, ScalarError> {
    if value.real == 0.0 && value.imaginary == 0.0 {
        return Err(ScalarError::Number);
    }
    let scale = value.real.abs().max(value.imaginary.abs());
    let ratio = value.real.abs().min(value.imaginary.abs()) / scale;
    let logarithm = scale.ln() + 0.5 * (ratio * ratio).ln_1p();
    finite(
        logarithm,
        argument(value.real, value.imaginary),
        value.suffix,
    )
}

/// Principal complex square root.  The component scale is factored out so
/// `f64::MAX + f64::MAX*i` remains usable even though its modulus overflows.
/// On the negative real axis, both signed-zero imaginary inputs use the
/// positive principal branch, consistent with the `(-pi, pi]` argument range.
pub(super) fn sqrt(value: Complex) -> Result<Complex, ScalarError> {
    if value.real == 0.0 && value.imaginary == 0.0 {
        return finite(0.0, value.imaginary, value.suffix);
    }

    let absolute_real = value.real.abs();
    let absolute_imaginary = value.imaginary.abs();
    let scale = absolute_real.max(absolute_imaginary);
    let real_component = absolute_real / scale;
    let imaginary_component = absolute_imaginary / scale;
    let normalized_magnitude = real_component.hypot(imaginary_component);
    let root_scale = scale.sqrt();
    let large_component = root_scale * (((normalized_magnitude + real_component) * 0.5).sqrt());

    if value.real >= 0.0 {
        let imaginary_magnitude = if absolute_imaginary == 0.0 {
            0.0
        } else {
            absolute_imaginary / (2.0 * large_component)
        };
        return finite(
            large_component,
            imaginary_magnitude.copysign(value.imaginary),
            value.suffix,
        );
    }

    let imaginary_magnitude = large_component;
    let real = if absolute_imaginary == 0.0 {
        0.0
    } else {
        absolute_imaginary / (2.0 * imaginary_magnitude)
    };
    let imaginary = imaginary_magnitude.copysign(if value.imaginary == 0.0 {
        1.0
    } else {
        value.imaginary
    });
    finite(real, imaginary, value.suffix)
}

/// Overflow-safe tangent.  For a large imaginary component the quotient is
/// expressed in terms of `exp(-2*abs(y))`, so both hyperbolic factors never
/// need to be materialized as infinities.
pub(super) fn tan(value: Complex) -> Result<Complex, ScalarError> {
    let y = value.imaginary.abs();
    if y < RATIO_TRIG_THRESHOLD {
        return div(sine(value)?, cosine(value)?);
    }
    let (sin_x, cos_x) = value.real.sin_cos();
    let (sin_twice_x, cos_twice_x) = double_angle(sin_x, cos_x);
    let q = (-2.0 * y).exp();
    let denominator = ratio_denominator(cos_twice_x, q, true);
    if denominator == 0.0 {
        return Err(ScalarError::DivisionByZero);
    }
    let sign = if value.imaginary.is_sign_negative() {
        -1.0
    } else {
        1.0
    };
    finite(
        2.0 * sin_twice_x * q / denominator,
        sign * (1.0 - q * q) / denominator,
        value.suffix,
    )
}

/// Overflow-safe cotangent.
pub(super) fn cot(value: Complex) -> Result<Complex, ScalarError> {
    let y = value.imaginary.abs();
    if y < RATIO_TRIG_THRESHOLD {
        return div(cosine(value)?, sine(value)?);
    }
    let (sin_x, cos_x) = value.real.sin_cos();
    let (sin_twice_x, cos_twice_x) = double_angle(sin_x, cos_x);
    let q = (-2.0 * y).exp();
    let denominator = ratio_denominator(cos_twice_x, q, false);
    if denominator == 0.0 {
        return Err(ScalarError::DivisionByZero);
    }
    let sign = if value.imaginary.is_sign_negative() {
        -1.0
    } else {
        1.0
    };
    finite(
        2.0 * sin_twice_x * q / denominator,
        -sign * (1.0 - q * q) / denominator,
        value.suffix,
    )
}

/// Overflow-safe secant.
pub(super) fn sec(value: Complex) -> Result<Complex, ScalarError> {
    let y = value.imaginary.abs();
    if y < RATIO_TRIG_THRESHOLD {
        return div(one(value.suffix), cosine(value)?);
    }
    let (sin_x, cos_x) = value.real.sin_cos();
    let (_, cos_twice_x) = double_angle(sin_x, cos_x);
    let q = (-2.0 * y).exp();
    let e = (-y).exp();
    let denominator = ratio_denominator(cos_twice_x, q, true);
    if denominator == 0.0 {
        return Err(ScalarError::DivisionByZero);
    }
    let sign = if value.imaginary.is_sign_negative() {
        -1.0
    } else {
        1.0
    };
    finite(
        2.0 * cos_x * e * (1.0 + q) / denominator,
        sign * 2.0 * sin_x * e * (1.0 - q) / denominator,
        value.suffix,
    )
}

/// Overflow-safe cosecant.
pub(super) fn csc(value: Complex) -> Result<Complex, ScalarError> {
    let y = value.imaginary.abs();
    if y < RATIO_TRIG_THRESHOLD {
        return div(one(value.suffix), sine(value)?);
    }
    let (sin_x, cos_x) = value.real.sin_cos();
    let (_, cos_twice_x) = double_angle(sin_x, cos_x);
    let q = (-2.0 * y).exp();
    let e = (-y).exp();
    let denominator = ratio_denominator(cos_twice_x, q, false);
    if denominator == 0.0 {
        return Err(ScalarError::DivisionByZero);
    }
    let sign = if value.imaginary.is_sign_negative() {
        -1.0
    } else {
        1.0
    };
    finite(
        2.0 * sin_x * e * (1.0 + q) / denominator,
        -sign * 2.0 * cos_x * e * (1.0 - q) / denominator,
        value.suffix,
    )
}

/// Overflow-safe hyperbolic secant.
pub(super) fn sech(value: Complex) -> Result<Complex, ScalarError> {
    let x = value.real.abs();
    if x < RATIO_TRIG_THRESHOLD {
        return div(one(value.suffix), hyperbolic_cosine(value)?);
    }
    let (sin_y, cos_y) = value.imaginary.sin_cos();
    let (_, cos_twice_y) = double_angle(sin_y, cos_y);
    let q = (-2.0 * x).exp();
    let e = (-x).exp();
    let denominator = ratio_denominator(cos_twice_y, q, true);
    if denominator == 0.0 {
        return Err(ScalarError::DivisionByZero);
    }
    let sign = if value.real.is_sign_negative() {
        -1.0
    } else {
        1.0
    };
    finite(
        2.0 * cos_y * e * (1.0 + q) / denominator,
        -sign * 2.0 * sin_y * e * (1.0 - q) / denominator,
        value.suffix,
    )
}

/// Overflow-safe hyperbolic cosecant.
pub(super) fn csch(value: Complex) -> Result<Complex, ScalarError> {
    let x = value.real.abs();
    if x < RATIO_TRIG_THRESHOLD {
        return div(one(value.suffix), hyperbolic_sine(value)?);
    }
    let (sin_y, cos_y) = value.imaginary.sin_cos();
    let (_, cos_twice_y) = double_angle(sin_y, cos_y);
    let q = (-2.0 * x).exp();
    let e = (-x).exp();
    let denominator = ratio_denominator(cos_twice_y, q, false);
    if denominator == 0.0 {
        return Err(ScalarError::DivisionByZero);
    }
    let sign = if value.real.is_sign_negative() {
        -1.0
    } else {
        1.0
    };
    finite(
        sign * 2.0 * cos_y * e * (1.0 - q) / denominator,
        -2.0 * sin_y * e * (1.0 + q) / denominator,
        value.suffix,
    )
}

fn one(suffix: char) -> Complex {
    Complex {
        real: 1.0,
        imaginary: 0.0,
        suffix,
    }
}

fn choose_suffix(left: char, right: char) -> char {
    if left == 'j' || right == 'j' {
        'j'
    } else {
        'i'
    }
}

fn argument(real: f64, imaginary: f64) -> f64 {
    let mut angle = imaginary.atan2(real);
    if angle == -std::f64::consts::PI {
        angle = std::f64::consts::PI;
    } else if angle == 0.0 {
        // Normalize both signed zeros on the positive real axis.
        angle = 0.0;
    }
    angle
}

fn double_angle(sine: f64, cosine: f64) -> (f64, f64) {
    (2.0 * sine * cosine, cosine.mul_add(cosine, -(sine * sine)))
}

fn ratio_denominator(cosine_twice: f64, q: f64, plus: bool) -> f64 {
    let sign = if plus { 1.0 } else { -1.0 };
    1.0 + sign * 2.0 * cosine_twice * q + q * q
}

#[derive(Clone, Copy)]
struct BinaryTerm {
    mantissa: f64,
    exponent: i32,
    negative: bool,
}

fn product_term(first: f64, second: f64) -> Result<Option<BinaryTerm>, ScalarError> {
    if !first.is_finite() || !second.is_finite() {
        return Err(ScalarError::Number);
    }
    if first == 0.0 || second == 0.0 {
        return Ok(None);
    }
    let (first_mantissa, first_exponent) = binary_parts(first.abs());
    let (second_mantissa, second_exponent) = binary_parts(second.abs());
    let (mantissa, exponent) = binary_parts(first_mantissa * second_mantissa);
    Ok(Some(BinaryTerm {
        mantissa,
        exponent: first_exponent + second_exponent + exponent,
        negative: first.is_sign_negative() ^ second.is_sign_negative(),
    }))
}

fn sum_products(
    first: Option<BinaryTerm>,
    second: Option<BinaryTerm>,
    subtract_second: bool,
) -> f64 {
    combine_terms(first, second, subtract_second)
        .map(scale_term)
        .unwrap_or(0.0)
}

fn combine_terms(
    first: Option<BinaryTerm>,
    second: Option<BinaryTerm>,
    subtract_second: bool,
) -> Option<BinaryTerm> {
    let Some(first) = first else {
        let mut second = second?;
        second.negative ^= subtract_second;
        return Some(second);
    };
    let Some(mut second) = second else {
        return Some(first);
    };
    second.negative ^= subtract_second;

    let first_is_larger = first.exponent > second.exponent
        || (first.exponent == second.exponent && first.mantissa >= second.mantissa);
    let (larger, smaller) = if first_is_larger {
        (first, second)
    } else {
        (second, first)
    };
    let exponent_gap = larger.exponent - smaller.exponent;
    let smaller_mantissa = scale_binary(smaller.mantissa, -exponent_gap);
    let combined = if larger.negative == smaller.negative {
        larger.mantissa + smaller_mantissa
    } else {
        larger.mantissa - smaller_mantissa
    };
    if combined == 0.0 {
        return None;
    }
    let negative = if combined.is_sign_negative() {
        !larger.negative
    } else {
        larger.negative
    };
    let (mantissa, exponent) = binary_parts(combined.abs());
    Some(BinaryTerm {
        mantissa,
        exponent: larger.exponent + exponent,
        negative,
    })
}

fn scale_term(term: BinaryTerm) -> f64 {
    scale_binary(
        term.mantissa
            .copysign(if term.negative { -1.0 } else { 1.0 }),
        term.exponent,
    )
}

fn divide_term(numerator: Option<BinaryTerm>, denominator: BinaryTerm) -> f64 {
    let Some(numerator) = numerator else {
        return 0.0;
    };
    let normalized = numerator.mantissa / denominator.mantissa;
    let (mantissa, exponent) = binary_parts(normalized);
    scale_binary(
        mantissa.copysign(if numerator.negative ^ denominator.negative {
            -1.0
        } else {
            1.0
        }),
        numerator.exponent - denominator.exponent + exponent,
    )
}

fn exponential_product(log_scale: f64, coefficient: f64) -> f64 {
    if coefficient == 0.0 {
        return coefficient;
    }
    if !coefficient.is_finite() || log_scale.is_nan() {
        return f64::NAN;
    }
    if log_scale <= 40.0 {
        return coefficient * log_scale.exp();
    }

    let log_absolute = log_scale + coefficient.abs().ln();
    scaled_exponential(log_absolute, coefficient)
}

fn scaled_exponential(log_absolute: f64, sign_source: f64) -> f64 {
    if log_absolute.is_nan() {
        return f64::NAN;
    }
    if log_absolute == f64::INFINITY {
        return f64::INFINITY.copysign(sign_source);
    }
    if log_absolute == f64::NEG_INFINITY {
        return 0.0_f64.copysign(sign_source);
    }

    let minimum_log = f64::from_bits(1).ln();
    if log_absolute < minimum_log {
        return 0.0_f64.copysign(sign_source);
    }
    if log_absolute > f64::MAX.ln() {
        return f64::INFINITY.copysign(sign_source);
    }

    let binary_exponent = (log_absolute / std::f64::consts::LN_2).floor();
    let fraction = log_absolute - binary_exponent * std::f64::consts::LN_2;
    scale_binary(fraction.exp().copysign(sign_source), binary_exponent as i32)
}

fn logarithm_cosh(value: f64) -> f64 {
    let value = value.abs();
    if value < RATIO_TRIG_THRESHOLD {
        value.cosh().ln()
    } else {
        value - std::f64::consts::LN_2 + (-2.0 * value).exp().ln_1p()
    }
}

fn logarithm_sinh_abs(value: f64) -> Option<f64> {
    let value = value.abs();
    if value == 0.0 {
        return None;
    }
    Some(if value < RATIO_TRIG_THRESHOLD {
        value.sinh().ln()
    } else {
        value - std::f64::consts::LN_2 + (-(-2.0 * value).exp()).ln_1p()
    })
}

fn hyperbolic_cosine_product(value: f64, coefficient: f64) -> f64 {
    exponential_product(logarithm_cosh(value), coefficient)
}

fn hyperbolic_sine_product(value: f64, coefficient: f64) -> f64 {
    let negative = coefficient.is_sign_negative() ^ value.is_sign_negative();
    let Some(logarithm) = logarithm_sinh_abs(value) else {
        return signed_zero(if negative { -1.0 } else { 1.0 });
    };
    exponential_product(
        logarithm,
        coefficient.copysign(if negative { -1.0 } else { 1.0 }),
    )
}

fn signed_zero(value: f64) -> f64 {
    0.0_f64.copysign(value)
}

fn sine(value: Complex) -> Result<Complex, ScalarError> {
    let (sin_real, cos_real) = value.real.sin_cos();
    finite(
        sin_real * value.imaginary.cosh(),
        cos_real * value.imaginary.sinh(),
        value.suffix,
    )
}

fn cosine(value: Complex) -> Result<Complex, ScalarError> {
    let (sin_real, cos_real) = value.real.sin_cos();
    finite(
        cos_real * value.imaginary.cosh(),
        -(sin_real * value.imaginary.sinh()),
        value.suffix,
    )
}

fn hyperbolic_sine(value: Complex) -> Result<Complex, ScalarError> {
    let (sin_imaginary, cos_imaginary) = value.imaginary.sin_cos();
    finite(
        value.real.sinh() * cos_imaginary,
        value.real.cosh() * sin_imaginary,
        value.suffix,
    )
}

fn hyperbolic_cosine(value: Complex) -> Result<Complex, ScalarError> {
    let (sin_imaginary, cos_imaginary) = value.imaginary.sin_cos();
    finite(
        value.real.cosh() * cos_imaginary,
        value.real.sinh() * sin_imaginary,
        value.suffix,
    )
}

fn binary_parts(value: f64) -> (f64, i32) {
    debug_assert!(value.is_finite() && value > 0.0);
    let bits = value.to_bits();
    let raw_exponent = ((bits >> 52) & 0x7ff) as i32;
    if raw_exponent == 0 {
        let (mantissa, exponent) = binary_parts(value * 18_014_398_509_481_984.0);
        return (mantissa, exponent - 54);
    }
    let fraction = bits & ((1_u64 << 52) - 1);
    let mantissa = f64::from_bits((1023_u64 << 52) | fraction);
    (mantissa, raw_exponent - 1023)
}

fn scale_binary(value: f64, exponent: i32) -> f64 {
    debug_assert!(value.is_finite() && value.abs() >= 1.0 && value.abs() < 2.0);
    let bits = value.to_bits();
    let sign = bits & (1_u64 << 63);
    let fraction = bits & ((1_u64 << 52) - 1);
    let target_exponent = 1023_i32 + exponent;
    if target_exponent >= 0x7ff {
        return f64::from_bits(sign | (0x7ff_u64 << 52));
    }
    if target_exponent >= 1 {
        return f64::from_bits(sign | ((target_exponent as u64) << 52) | fraction);
    }

    let significand = (1_u64 << 52) | fraction;
    let right_shift = 1_i32 - target_exponent;
    if right_shift >= 64 {
        return f64::from_bits(sign);
    }
    let right_shift = right_shift as u32;
    let mask = (1_u64 << right_shift) - 1;
    let mut subnormal = significand >> right_shift;
    let remainder = significand & mask;
    let halfway = 1_u64 << (right_shift - 1);
    if remainder > halfway || (remainder == halfway && subnormal & 1 != 0) {
        subnormal += 1;
    }
    if subnormal >= 1_u64 << 52 {
        return f64::from_bits(sign | (1_u64 << 52));
    }
    f64::from_bits(sign | subnormal)
}
