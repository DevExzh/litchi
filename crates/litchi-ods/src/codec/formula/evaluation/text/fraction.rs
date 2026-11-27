//! Exact bounded rational approximation for TEXT's simple fractions.
//!
//! The input is the exact binary64 value supplied by the caller.  The helper
//! compares rational distances with integer cross products, so decimal
//! conversion and a million-candidate floating-point scan cannot change the
//! selected fraction.  Work and cancellation remain caller-owned through the
//! callback.

use super::super::ScalarError;

const MAX_DENOMINATOR: u64 = 999_999;

/// Find the closest non-negative rational with denominator at most
/// `max_denominator`.
///
/// The outer result carries the caller's work/cancellation failure.  The
/// inner result carries the formula-level numeric failure.  The returned
/// fraction is reduced, except that `(1, 1)` is retained so the formatter can
/// carry a rounded fractional part into its whole-number field.  Equal
/// distances choose the smaller denominator, then the smaller numerator.
pub(super) fn nearest<E, F>(
    value: f64,
    max_denominator: u64,
    mut charge: F,
) -> Result<Result<(u64, u64), ScalarError>, E>
where
    F: FnMut(u64) -> Result<(), E>,
{
    if !value.is_finite()
        || !(0.0..=1.0).contains(&value)
        || !(1..=MAX_DENOMINATOR).contains(&max_denominator)
    {
        return Ok(Err(ScalarError::Number));
    }

    let (numerator, shift) = decode_dyadic(value);
    if numerator == 0 {
        return Ok(Ok((0, 1)));
    }
    // For Q <= 999999, every value with denominator 2^74 or larger is below
    // half of 1/Q.  It therefore rounds exactly to zero; this also covers all
    // positive subnormals without constructing a 1074-bit denominator.
    if shift > 73 {
        return Ok(Ok((0, 1)));
    }
    let denominator = match 1_u128.checked_shl(shift) {
        Some(value) => value,
        None => return Ok(Err(ScalarError::Number)),
    };
    nearest_dyadic(
        numerator,
        denominator,
        u128::from(max_denominator),
        &mut charge,
    )
}

fn decode_dyadic(value: f64) -> (u128, u32) {
    if value == 0.0 {
        return (0, 0);
    }
    let bits = value.to_bits();
    let exponent = (bits >> 52) & 0x7ff;
    let fraction = bits & ((1_u64 << 52) - 1);
    if exponent == 0 {
        (u128::from(fraction), 1_074)
    } else {
        // The caller has already restricted the value to [0, 1], so a
        // nonzero exponent is in 1..=1023.  The casts therefore stay within
        // the binary64 exponent range and produce a shift in 52..=1073.
        let exponent = exponent as i32 - 1_023;
        let shift = (52_i32 - exponent) as u32;
        (u128::from(fraction | (1_u64 << 52)), shift)
    }
}

#[derive(Clone, Copy)]
struct Candidate {
    numerator: u128,
    denominator: u128,
    error: u128,
}

fn nearest_dyadic<E, F>(
    numerator: u128,
    denominator: u128,
    max_denominator: u128,
    charge: &mut F,
) -> Result<Result<(u64, u64), ScalarError>, E>
where
    F: FnMut(u64) -> Result<(), E>,
{
    let divisor = gcd(numerator, denominator);
    let numerator = numerator / divisor;
    let denominator = denominator / divisor;
    let mut best = Candidate {
        numerator: 0,
        denominator: 1,
        error: numerator,
    };

    // Convergents of the continued fraction are the only candidates needed,
    // together with the final bounded semiconvergent.  All state is u128:
    // the decoded denominator is at most 2^73, while Q is at most 999999.
    let mut p_before_previous = 0_u128;
    let mut p_previous = 1_u128;
    let mut q_before_previous = 1_u128;
    let mut q_previous = 0_u128;
    let mut remainder_numerator = numerator;
    let mut remainder_denominator = denominator;

    while remainder_denominator != 0 {
        charge(1)?;
        let quotient = remainder_numerator / remainder_denominator;
        let product = match quotient.checked_mul(p_previous) {
            Some(value) => value,
            None => return Ok(Err(ScalarError::Number)),
        };
        let candidate_numerator = match p_before_previous.checked_add(product) {
            Some(value) => value,
            None => return Ok(Err(ScalarError::Number)),
        };
        let product = match quotient.checked_mul(q_previous) {
            Some(value) => value,
            None => return Ok(Err(ScalarError::Number)),
        };
        let candidate_denominator = match q_before_previous.checked_add(product) {
            Some(value) => value,
            None => return Ok(Err(ScalarError::Number)),
        };

        if candidate_denominator > max_denominator {
            if q_previous == 0 {
                return Ok(Err(ScalarError::Number));
            }
            let remaining = match max_denominator.checked_sub(q_before_previous) {
                Some(value) => value,
                None => return Ok(Err(ScalarError::Number)),
            };
            let multiple = remaining / q_previous;
            if multiple != 0 {
                let product = match multiple.checked_mul(p_previous) {
                    Some(value) => value,
                    None => return Ok(Err(ScalarError::Number)),
                };
                let semiconvergent_numerator = match p_before_previous.checked_add(product) {
                    Some(value) => value,
                    None => return Ok(Err(ScalarError::Number)),
                };
                let product = match multiple.checked_mul(q_previous) {
                    Some(value) => value,
                    None => return Ok(Err(ScalarError::Number)),
                };
                let semiconvergent_denominator = match q_before_previous.checked_add(product) {
                    Some(value) => value,
                    None => return Ok(Err(ScalarError::Number)),
                };
                let candidate = match make_candidate(
                    numerator,
                    denominator,
                    semiconvergent_numerator,
                    semiconvergent_denominator,
                ) {
                    Ok(value) => value,
                    Err(error) => return Ok(Err(error)),
                };
                if let Err(error) = update_best(&mut best, candidate) {
                    return Ok(Err(error));
                }
            }
            break;
        }

        let candidate = match make_candidate(
            numerator,
            denominator,
            candidate_numerator,
            candidate_denominator,
        ) {
            Ok(value) => value,
            Err(error) => return Ok(Err(error)),
        };
        if let Err(error) = update_best(&mut best, candidate) {
            return Ok(Err(error));
        }

        p_before_previous = p_previous;
        p_previous = candidate_numerator;
        q_before_previous = q_previous;
        q_previous = candidate_denominator;
        let next_remainder = remainder_numerator % remainder_denominator;
        remainder_numerator = remainder_denominator;
        remainder_denominator = next_remainder;
    }

    let numerator = match u64::try_from(best.numerator) {
        Ok(value) => value,
        Err(_) => return Ok(Err(ScalarError::Number)),
    };
    let denominator = match u64::try_from(best.denominator) {
        Ok(value) => value,
        Err(_) => return Ok(Err(ScalarError::Number)),
    };
    let divisor = gcd_u64(numerator, denominator);
    Ok(Ok((numerator / divisor, denominator / divisor)))
}

fn make_candidate(
    numerator: u128,
    denominator: u128,
    candidate_numerator: u128,
    candidate_denominator: u128,
) -> Result<Candidate, ScalarError> {
    if candidate_denominator == 0 {
        return Err(ScalarError::Number);
    }
    let left = numerator
        .checked_mul(candidate_denominator)
        .ok_or(ScalarError::Number)?;
    let right = candidate_numerator
        .checked_mul(denominator)
        .ok_or(ScalarError::Number)?;
    Ok(Candidate {
        numerator: candidate_numerator,
        denominator: candidate_denominator,
        error: left.abs_diff(right),
    })
}

fn update_best(best: &mut Candidate, candidate: Candidate) -> Result<(), ScalarError> {
    if is_better(candidate, *best)? {
        *best = candidate;
    }
    Ok(())
}

fn is_better(candidate: Candidate, current: Candidate) -> Result<bool, ScalarError> {
    let left = candidate
        .error
        .checked_mul(current.denominator)
        .ok_or(ScalarError::Number)?;
    let right = current
        .error
        .checked_mul(candidate.denominator)
        .ok_or(ScalarError::Number)?;
    Ok(left < right
        || (left == right
            && (candidate.denominator < current.denominator
                || (candidate.denominator == current.denominator
                    && candidate.numerator < current.numerator))))
}

fn gcd(mut left: u128, mut right: u128) -> u128 {
    while right != 0 {
        let remainder = left % right;
        left = right;
        right = remainder;
    }
    left
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
    use super::{MAX_DENOMINATOR, nearest};
    use crate::codec::formula::evaluation::ScalarError;

    fn exact_dyadic(value: f64) -> (u128, u32) {
        if value == 0.0 {
            return (0, 0);
        }
        let bits = value.to_bits();
        let exponent = (bits >> 52) & 0x7ff;
        let fraction = bits & ((1_u64 << 52) - 1);
        if exponent == 0 {
            (u128::from(fraction), 1_074)
        } else {
            let exponent = i32::try_from(exponent).expect("test exponent") - 1_023;
            let shift = u32::try_from(52_i32 - exponent).expect("test shift");
            (u128::from(fraction | (1_u64 << 52)), shift)
        }
    }

    fn reference(value: f64, max_denominator: u64) -> (u64, u64) {
        let (numerator, shift) = exact_dyadic(value);
        let denominator = 1_u128.checked_shl(shift).expect("small test dyadic");
        let mut best = (0_u64, 1_u64);
        for candidate_denominator in 1..=max_denominator {
            for candidate_numerator in 0..=candidate_denominator {
                let candidate_error = numerator
                    .checked_mul(u128::from(candidate_denominator))
                    .expect("test product")
                    .abs_diff(
                        u128::from(candidate_numerator)
                            .checked_mul(denominator)
                            .expect("test product"),
                    );
                let best_error = numerator
                    .checked_mul(u128::from(best.1))
                    .expect("test product")
                    .abs_diff(
                        u128::from(best.0)
                            .checked_mul(denominator)
                            .expect("test product"),
                    );
                let left = candidate_error * u128::from(best.1);
                let right = best_error * u128::from(candidate_denominator);
                if left < right
                    || (left == right
                        && (candidate_denominator < best.1
                            || (candidate_denominator == best.1 && candidate_numerator < best.0)))
                {
                    best = (candidate_numerator, candidate_denominator);
                }
            }
        }
        let divisor = gcd_test(best.0, best.1);
        (best.0 / divisor, best.1 / divisor)
    }

    fn gcd_test(mut left: u64, mut right: u64) -> u64 {
        while right != 0 {
            let remainder = left % right;
            left = right;
            right = remainder;
        }
        left
    }

    fn assert_fraction(value: f64, max_denominator: u64, expected: (u64, u64)) {
        let actual = nearest(value, max_denominator, |_amount| Ok::<(), ()>(()))
            .expect("test charge")
            .expect("valid fraction");
        assert_eq!(actual, expected, "value={value:?} max={max_denominator}");
    }

    #[test]
    fn exhaustive_small_dyadics_match_integer_reference() {
        for shift in 0..=8_u32 {
            let denominator = 1_u128 << shift;
            for numerator in 0..=denominator {
                let value = (numerator as f64) / (denominator as f64);
                for max_denominator in 1..=20_u64 {
                    let expected = reference(value, max_denominator);
                    assert_fraction(value, max_denominator, expected);
                }
            }
        }
    }

    #[test]
    fn reported_boundary_regressions_use_exact_distances() {
        assert_fraction(0.11805555555555555, 9, (1, 9));
        assert_fraction(0.13392857142857142, 8, (1, 8));
        assert_fraction(0.2361111111111111, 9, (2, 9));
        assert_fraction(0.005050505050505051, 99, (1, 99));
    }

    #[test]
    fn zero_one_subnormal_and_carry_are_stable() {
        assert_fraction(0.0, 1, (0, 1));
        assert_fraction(1.0, MAX_DENOMINATOR, (1, 1));
        assert_fraction(f64::from_bits(1), MAX_DENOMINATOR, (0, 1));
        assert_fraction(0.5, 1, (0, 1));
        assert_fraction(0.5, 2, (1, 2));
    }

    #[test]
    fn invalid_inputs_are_formula_number_errors() {
        for (value, max_denominator) in [
            (f64::NAN, 9),
            (-0.1, 9),
            (1.1, 9),
            (0.5, 0),
            (0.5, MAX_DENOMINATOR + 1),
        ] {
            let result =
                nearest(value, max_denominator, |_amount| Ok::<(), ()>(())).expect("test charge");
            assert_eq!(result, Err(ScalarError::Number));
        }
    }

    #[test]
    fn work_failure_propagates_before_later_iterations() {
        let mut charges = 0_u64;
        let result = nearest(0.11805555555555555, 999_999, |amount| {
            charges += amount;
            if charges > 2 {
                Err("cancelled")
            } else {
                Ok(())
            }
        });
        assert_eq!(result, Err("cancelled"));
        assert!(charges <= 3);
    }
}
