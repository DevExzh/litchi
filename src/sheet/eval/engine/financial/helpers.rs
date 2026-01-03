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
