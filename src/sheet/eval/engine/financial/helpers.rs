use crate::sheet::Result;
use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;

pub(super) async fn number_arg(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    expr: &Expr,
) -> Result<Option<f64>> {
    let v = evaluate_expression(ctx, current_sheet, expr).await?;
    Ok(to_number(&v))
}

pub(super) fn present_value(rate: f64, nper: f64, pmt: f64, fv: f64, typ: f64) -> f64 {
    if rate.abs() < 1e-10 {
        -(pmt * nper + fv)
    } else {
        let r1 = 1.0 + rate;
        let factor = r1.powf(-nper);
        let pv_pmt = pmt * (1.0 + rate * typ) * (1.0 - factor) / rate;
        let pv_fv = fv * factor;
        -(pv_pmt + pv_fv)
    }
}

pub(super) fn future_value(rate: f64, nper: f64, pmt: f64, pv: f64, typ: f64) -> f64 {
    if rate.abs() < 1e-10 {
        -(pv + pmt * nper)
    } else {
        let r1 = 1.0 + rate;
        let factor = r1.powf(nper);
        let fv_pv = pv * factor;
        let fv_pmt = pmt * (1.0 + rate * typ) * (factor - 1.0) / rate;
        -(fv_pv + fv_pmt)
    }
}

pub(super) fn solve_rate(
    nper: f64,
    pmt: f64,
    pv: f64,
    fv: f64,
    typ: f64,
    guess: f64,
) -> Option<f64> {
    let mut rate = guess;
    let max_iter = 100;
    let tol = 1e-10;

    for _ in 0..max_iter {
        let f = rate_function(rate, nper, pmt, pv, fv, typ);
        if f.abs() < tol {
            return Some(rate);
        }
        let deriv = numerical_derivative(|r| rate_function(r, nper, pmt, pv, fv, typ), rate);
        if deriv.abs() < 1e-12 {
            break;
        }
        let new_rate = rate - f / deriv;
        if !new_rate.is_finite() {
            break;
        }
        rate = new_rate;
    }

    None
}

pub(super) fn solve_irr(cash_flows: &[f64], guess: f64) -> Option<f64> {
    let mut rate = guess;
    let max_iter = 100;
    let tol = 1e-10;

    for _ in 0..max_iter {
        let f = npv_for_irr(rate, cash_flows);
        if f.abs() < tol {
            return Some(rate);
        }
        let deriv = numerical_derivative(|r| npv_for_irr(r, cash_flows), rate);
        if deriv.abs() < 1e-12 {
            break;
        }
        let new_rate = rate - f / deriv;
        if !new_rate.is_finite() {
            break;
        }
        rate = new_rate;
    }

    None
}

pub(super) fn xnpv(rate: f64, cash_flows: &[f64], dates: &[f64]) -> f64 {
    let base_date = dates[0];
    let mut total = 0.0;
    for (cf, d) in cash_flows.iter().zip(dates.iter()) {
        let t = (d - base_date) / 365.0;
        let denom = (1.0 + rate).powf(t);
        total += cf / denom;
    }
    total
}

pub(super) fn solve_xirr(cash_flows: &[f64], dates: &[f64], guess: f64) -> Option<f64> {
    let mut rate = guess;
    let max_iter = 100;
    let tol = 1e-10;

    for _ in 0..max_iter {
        let f = xnpv(rate, cash_flows, dates);
        if f.abs() < tol {
            return Some(rate);
        }
        let deriv = numerical_derivative(|r| xnpv(r, cash_flows, dates), rate);
        if deriv.abs() < 1e-12 {
            break;
        }
        let new_rate = rate - f / deriv;
        if !new_rate.is_finite() {
            break;
        }
        rate = new_rate;
    }

    None
}

fn rate_function(rate: f64, nper: f64, pmt: f64, pv: f64, fv: f64, typ: f64) -> f64 {
    if rate.abs() < 1e-10 {
        pv + pmt * nper + fv
    } else {
        let r1 = 1.0 + rate;
        let factor = r1.powf(-nper);
        let term1 = pmt * (1.0 + rate * typ) * (1.0 - factor) / rate;
        let term2 = fv * factor;
        pv + term1 + term2
    }
}

fn npv_for_irr(rate: f64, cash_flows: &[f64]) -> f64 {
    let mut total = 0.0;
    for (i, cf) in cash_flows.iter().enumerate() {
        let t = i as f64;
        let denom = (1.0 + rate).powf(t);
        total += cf / denom;
    }
    total
}

fn numerical_derivative<F>(f: F, x: f64) -> f64
where
    F: Fn(f64) -> f64,
{
    let h = 1e-5;
    let f1 = f(x + h);
    let f2 = f(x - h);
    (f1 - f2) / (2.0 * h)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_present_value_zero_rate() {
        // PV with zero rate: -(pmt * nper + fv)
        let pv = present_value(0.0, 10.0, -100.0, 1000.0, 0.0);
        // -( -100 * 10 + 1000 ) = -( -1000 + 1000 ) = 0
        assert!((pv - 0.0).abs() < 1e-9);
    }

    #[test]
    fn test_present_value_with_rate() {
        // PV of annuity: 10 periods, $100 payment, 5% rate
        let pv = present_value(0.05, 10.0, -100.0, 0.0, 0.0);
        // Expected: 100 * (1 - 1.05^-10) / 0.05 ≈ 772.17
        assert!((pv - 772.17).abs() < 1.0);
    }

    #[test]
    fn test_present_value_with_fv() {
        // PV with future value
        let pv = present_value(0.05, 10.0, 0.0, 1000.0, 0.0);
        // Expected: -1000 / 1.05^10 ≈ -613.91
        assert!((pv - (-613.91)).abs() < 1.0);
    }

    #[test]
    fn test_present_value_with_type() {
        // PV with beginning-of-period payments (typ=1)
        let pv = present_value(0.05, 10.0, -100.0, 0.0, 1.0);
        // Expected: 100 * (1 + 0.05) * (1 - 1.05^-10) / 0.05 ≈ 810.78
        assert!((pv - 810.78).abs() < 1.0);
    }

    #[test]
    fn test_future_value_zero_rate() {
        // FV with zero rate: -(pv + pmt * nper)
        let fv = future_value(0.0, 10.0, -100.0, -1000.0, 0.0);
        // -( -1000 + (-100) * 10 ) = -( -1000 - 1000 ) = 2000
        assert!((fv - 2000.0).abs() < 1e-9);
    }

    #[test]
    fn test_future_value_with_rate() {
        // FV of annuity: 10 periods, $100 payment, 5% rate
        let fv = future_value(0.05, 10.0, -100.0, 0.0, 0.0);
        // Expected: 100 * (1.05^10 - 1) / 0.05 ≈ 1257.79
        assert!((fv - 1257.79).abs() < 1.0);
    }

    #[test]
    fn test_future_value_with_pv() {
        // FV with present value
        let fv = future_value(0.05, 10.0, 0.0, -1000.0, 0.0);
        // Expected: 1000 * 1.05^10 ≈ 1628.89
        assert!((fv - 1628.89).abs() < 1.0);
    }

    #[test]
    fn test_future_value_with_type() {
        // FV with beginning-of-period payments (typ=1)
        let fv = future_value(0.05, 10.0, -100.0, 0.0, 1.0);
        // Expected: 100 * (1 + 0.05) * (1.05^10 - 1) / 0.05 ≈ 1320.68
        assert!((fv - 1320.68).abs() < 1.0);
    }

    #[test]
    fn test_solve_rate_basic() {
        // Solve for rate given nper=10, pmt=-100, pv=772.17, fv=0
        let rate = solve_rate(10.0, -100.0, 772.17, 0.0, 0.0, 0.1);
        assert!(rate.is_some());
        let r = rate.unwrap();
        assert!((r - 0.05).abs() < 0.01);
    }

    #[test]
    fn test_solve_rate_with_fv() {
        // Solve for rate given nper=10, pmt=0, pv=-1000, fv=1628.89
        let rate = solve_rate(10.0, 0.0, -1000.0, 1628.89, 0.0, 0.1);
        assert!(rate.is_some());
        let r = rate.unwrap();
        assert!((r - 0.05).abs() < 0.01);
    }

    #[test]
    fn test_solve_irr_basic() {
        // IRR for cash flows: -1000, 300, 400, 500
        // This should give a positive rate
        let cash_flows = vec![-1000.0, 300.0, 400.0, 500.0];
        let irr = solve_irr(&cash_flows, 0.1);
        assert!(irr.is_some());
        let r = irr.unwrap();
        assert!(r > 0.0 && r < 0.5);
    }

    #[test]
    fn test_solve_irr_simple() {
        // Simple IRR: -1000, 1100 after 1 period -> 10%
        let cash_flows = vec![-1000.0, 1100.0];
        let irr = solve_irr(&cash_flows, 0.1);
        assert!(irr.is_some());
        let r = irr.unwrap();
        assert!((r - 0.10).abs() < 0.01);
    }

    #[test]
    fn test_solve_irr_no_solution() {
        // All positive cash flows - no IRR
        let cash_flows = vec![1000.0, 100.0, 200.0];
        let _irr = solve_irr(&cash_flows, 0.1);
        // May converge to something or return None
        // Just ensure it doesn't panic
    }

    #[test]
    fn test_xnpv_basic() {
        // XNPV with dates 0, 365, 730 days
        let cash_flows = vec![-1000.0, 300.0, 800.0];
        let dates = vec![0.0, 365.0, 730.0];
        let npv = xnpv(0.05, &cash_flows, &dates);
        // NPV = -1000 + 300/1.05^1 + 800/1.05^2
        assert!(npv.is_finite());
        assert!(npv > -1000.0);
    }

    #[test]
    fn test_xnpv_zero_rate() {
        let cash_flows = vec![-1000.0, 300.0, 800.0];
        let dates = vec![0.0, 365.0, 730.0];
        let npv = xnpv(0.0, &cash_flows, &dates);
        // With zero rate, just sum of cash flows
        assert!((npv - 100.0).abs() < 1e-9);
    }

    #[test]
    fn test_solve_xirr_basic() {
        // XIRR for cash flows over 2 years
        let cash_flows = vec![-1000.0, 600.0, 600.0];
        let dates = vec![0.0, 365.0, 730.0];
        let xirr = solve_xirr(&cash_flows, &dates, 0.1);
        assert!(xirr.is_some());
        let r = xirr.unwrap();
        assert!(r > 0.0 && r < 0.5);
    }

    #[test]
    fn test_numerical_derivative_linear() {
        // Derivative of f(x) = 2x + 3 is 2
        let f = |x: f64| 2.0 * x + 3.0;
        let deriv = numerical_derivative(f, 1.0);
        assert!((deriv - 2.0).abs() < 1e-4);
    }

    #[test]
    fn test_numerical_derivative_quadratic() {
        // Derivative of f(x) = x^2 is 2x
        let f = |x: f64| x * x;
        let deriv_at_3 = numerical_derivative(f, 3.0);
        assert!((deriv_at_3 - 6.0).abs() < 1e-4);
    }

    #[test]
    fn test_numerical_derivative_exp() {
        // Derivative of f(x) = e^x is e^x
        let f = |x: f64| x.exp();
        let deriv_at_0 = numerical_derivative(f, 0.0);
        assert!((deriv_at_0 - 1.0).abs() < 1e-4);
    }
}
