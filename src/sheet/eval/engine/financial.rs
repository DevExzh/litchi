use crate::sheet::{CellValue, Result};

use super::super::parser::Expr;
use super::{EngineCtx, evaluate_expression, flatten_range_expr, to_number};

pub(crate) fn eval_pv<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 5 {
        return Ok(CellValue::Error(
            "PV expects 3 to 5 arguments (rate, nper, pmt, [fv], [type])".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("PV rate is not numeric".to_string())),
    };
    let nper = match number_arg(ctx, current_sheet, &args[1])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("PV nper is not numeric".to_string())),
    };
    let pmt = match number_arg(ctx, current_sheet, &args[2])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("PV pmt is not numeric".to_string())),
    };

    let fv = if args.len() >= 4 {
        match number_arg(ctx, current_sheet, &args[3])? {
            Some(v) => v,
            None => return Ok(CellValue::Error("PV fv is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let typ = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4])? {
            Some(v) => v,
            None => return Ok(CellValue::Error("PV type is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let pv = present_value(rate, nper, pmt, fv, typ);
    Ok(CellValue::Float(pv))
}

pub(crate) fn eval_fv<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 5 {
        return Ok(CellValue::Error(
            "FV expects 3 to 5 arguments (rate, nper, pmt, [pv], [type])".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("FV rate is not numeric".to_string())),
    };
    let nper = match number_arg(ctx, current_sheet, &args[1])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("FV nper is not numeric".to_string())),
    };
    let pmt = match number_arg(ctx, current_sheet, &args[2])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("FV pmt is not numeric".to_string())),
    };

    let pv = if args.len() >= 4 {
        match number_arg(ctx, current_sheet, &args[3])? {
            Some(v) => v,
            None => return Ok(CellValue::Error("FV pv is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let typ = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4])? {
            Some(v) => v,
            None => return Ok(CellValue::Error("FV type is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let fv = future_value(rate, nper, pmt, pv, typ);
    Ok(CellValue::Float(fv))
}

pub(crate) fn eval_rate<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 6 {
        return Ok(CellValue::Error(
            "RATE expects 3 to 6 arguments (nper, pmt, pv, [fv], [type], [guess])".to_string(),
        ));
    }

    let nper = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("RATE nper is not numeric".to_string())),
    };
    let pmt = match number_arg(ctx, current_sheet, &args[1])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("RATE pmt is not numeric".to_string())),
    };
    let pv = match number_arg(ctx, current_sheet, &args[2])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("RATE pv is not numeric".to_string())),
    };

    let fv = if args.len() >= 4 {
        match number_arg(ctx, current_sheet, &args[3])? {
            Some(v) => v,
            None => return Ok(CellValue::Error("RATE fv is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let typ = if args.len() >= 5 {
        match number_arg(ctx, current_sheet, &args[4])? {
            Some(v) => v,
            None => return Ok(CellValue::Error("RATE type is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let guess = if args.len() == 6 {
        match number_arg(ctx, current_sheet, &args[5])? {
            Some(v) => v,
            None => return Ok(CellValue::Error("RATE guess is not numeric".to_string())),
        }
    } else {
        0.1
    };

    let rate = solve_rate(nper, pmt, pv, fv, typ, guess);
    match rate {
        Some(r) => Ok(CellValue::Float(r)),
        None => Ok(CellValue::Error("RATE failed to converge".to_string())),
    }
}

pub(crate) fn eval_npv<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 {
        return Ok(CellValue::Error(
            "NPV expects at least 2 arguments (rate, value1, [value2], ...)".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("NPV rate is not numeric".to_string())),
    };

    let mut total = 0.0f64;
    let mut period = 1.0f64;

    for expr in &args[1..] {
        let range = flatten_range_expr(ctx, current_sheet, expr)?;
        for v in range.values {
            if let Some(cf) = to_number(&v) {
                let denom = (1.0 + rate).powf(period);
                total += cf / denom;
                period += 1.0;
            }
        }
    }

    Ok(CellValue::Float(total))
}

pub(crate) fn eval_irr<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "IRR expects 1 or 2 arguments (values, [guess])".to_string(),
        ));
    }

    let range = flatten_range_expr(ctx, current_sheet, &args[0])?;
    let mut cash_flows = Vec::new();
    for v in range.values {
        if let Some(n) = to_number(&v) {
            cash_flows.push(n);
        }
    }

    if cash_flows.is_empty() {
        return Ok(CellValue::Error(
            "IRR requires at least one cash flow".to_string(),
        ));
    }

    let has_pos = cash_flows.iter().any(|v| *v > 0.0);
    let has_neg = cash_flows.iter().any(|v| *v < 0.0);
    if !has_pos || !has_neg {
        return Ok(CellValue::Error(
            "IRR requires at least one positive and one negative cash flow".to_string(),
        ));
    }

    let guess = if args.len() == 2 {
        match number_arg(ctx, current_sheet, &args[1])? {
            Some(v) => v,
            None => return Ok(CellValue::Error("IRR guess is not numeric".to_string())),
        }
    } else {
        0.1
    };

    let irr = solve_irr(&cash_flows, guess);
    match irr {
        Some(r) => Ok(CellValue::Float(r)),
        None => Ok(CellValue::Error("IRR failed to converge".to_string())),
    }
}

pub(crate) fn eval_xnpv<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "XNPV expects 3 arguments (rate, values, dates)".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("XNPV rate is not numeric".to_string())),
    };

    let values_range = flatten_range_expr(ctx, current_sheet, &args[1])?;
    let dates_range = flatten_range_expr(ctx, current_sheet, &args[2])?;

    if values_range.values.len() != dates_range.values.len() || values_range.values.is_empty() {
        return Ok(CellValue::Error(
            "XNPV requires values and dates ranges of the same non-zero length".to_string(),
        ));
    }

    let mut cash_flows = Vec::new();
    let mut dates = Vec::new();

    for v in values_range.values {
        if let Some(n) = to_number(&v) {
            cash_flows.push(n);
        } else {
            cash_flows.push(0.0);
        }
    }

    for v in dates_range.values {
        if let Some(n) = to_number(&v) {
            dates.push(n);
        } else {
            return Ok(CellValue::Error(
                "XNPV dates must be numeric serials".to_string(),
            ));
        }
    }

    let npv = xnpv(rate, &cash_flows, &dates);
    Ok(CellValue::Float(npv))
}

pub(crate) fn eval_xirr<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "XIRR expects 2 or 3 arguments (values, dates, [guess])".to_string(),
        ));
    }

    let values_range = flatten_range_expr(ctx, current_sheet, &args[0])?;
    let dates_range = flatten_range_expr(ctx, current_sheet, &args[1])?;

    if values_range.values.len() != dates_range.values.len() || values_range.values.is_empty() {
        return Ok(CellValue::Error(
            "XIRR requires values and dates ranges of the same non-zero length".to_string(),
        ));
    }

    let mut cash_flows = Vec::new();
    let mut dates = Vec::new();

    for v in values_range.values {
        if let Some(n) = to_number(&v) {
            cash_flows.push(n);
        } else {
            cash_flows.push(0.0);
        }
    }

    for v in dates_range.values {
        if let Some(n) = to_number(&v) {
            dates.push(n);
        } else {
            return Ok(CellValue::Error(
                "XIRR dates must be numeric serials".to_string(),
            ));
        }
    }

    let has_pos = cash_flows.iter().any(|v| *v > 0.0);
    let has_neg = cash_flows.iter().any(|v| *v < 0.0);
    if !has_pos || !has_neg {
        return Ok(CellValue::Error(
            "XIRR requires at least one positive and one negative cash flow".to_string(),
        ));
    }

    let guess = if args.len() == 3 {
        match number_arg(ctx, current_sheet, &args[2])? {
            Some(v) => v,
            None => return Ok(CellValue::Error("XIRR guess is not numeric".to_string())),
        }
    } else {
        0.1
    };

    let irr = solve_xirr(&cash_flows, &dates, guess);
    match irr {
        Some(r) => Ok(CellValue::Float(r)),
        None => Ok(CellValue::Error("XIRR failed to converge".to_string())),
    }
}

pub(crate) fn eval_yield<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 6 || args.len() > 7 {
        return Ok(CellValue::Error(
            "YIELD expects 6 or 7 arguments (settlement, maturity, rate, pr, redemption, frequency, [basis])"
                .to_string(),
        ));
    }

    let settlement = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "YIELD settlement is not numeric".to_string(),
            ));
        },
    };
    let maturity = match number_arg(ctx, current_sheet, &args[1])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "YIELD maturity is not numeric".to_string(),
            ));
        },
    };
    if maturity <= settlement {
        return Ok(CellValue::Error(
            "YIELD requires maturity to be after settlement".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[2])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("YIELD rate is not numeric".to_string())),
    };
    let price = match number_arg(ctx, current_sheet, &args[3])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("YIELD pr is not numeric".to_string())),
    };
    let redemption = match number_arg(ctx, current_sheet, &args[4])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "YIELD redemption is not numeric".to_string(),
            ));
        },
    };
    let freq_val = match number_arg(ctx, current_sheet, &args[5])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "YIELD frequency is not numeric".to_string(),
            ));
        },
    };

    let freq = freq_val.trunc() as i32;
    if !matches!(freq, 1 | 2 | 4) {
        return Ok(CellValue::Error(
            "YIELD frequency must be 1 (annual), 2 (semiannual), or 4 (quarterly)".to_string(),
        ));
    }

    let basis = if args.len() == 7 {
        match number_arg(ctx, current_sheet, &args[6])? {
            Some(v) => v,
            None => return Ok(CellValue::Error("YIELD basis is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let basis_int = basis.trunc() as i32;
    if !(0..=4).contains(&basis_int) {
        return Ok(CellValue::Error(
            "YIELD basis must be between 0 and 4".to_string(),
        ));
    }

    let days = maturity - settlement;
    if days <= 0.0 {
        return Ok(CellValue::Error(
            "YIELD requires maturity to be after settlement".to_string(),
        ));
    }

    let year_days = match basis_int {
        0 | 2 => 360.0,
        _ => 365.0,
    };

    let years = days / year_days;
    let mut nper = (years * freq as f64).round() as i32;
    if nper <= 0 {
        nper = 1;
    }

    // Simple coupon bond model: coupons based on 100 par, paid "freq" times per year.
    let coupon = rate * 100.0 / freq as f64;

    let mut cash_flows = Vec::with_capacity(nper as usize + 1);
    cash_flows.push(-price);
    for _ in 1..nper {
        cash_flows.push(coupon);
    }
    cash_flows.push(coupon + redemption);

    let guess = rate / freq as f64;
    let per_period_yield = match solve_irr(&cash_flows, guess) {
        Some(r) => r,
        None => {
            return Ok(CellValue::Error("YIELD failed to converge".to_string()));
        },
    };

    let annual_yield = per_period_yield * freq as f64;
    Ok(CellValue::Float(annual_yield))
}

pub(crate) fn eval_duration<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 5 || args.len() > 6 {
        return Ok(CellValue::Error(
            "DURATION expects 5 or 6 arguments (settlement, maturity, coupon, yld, frequency, [basis])"
                .to_string(),
        ));
    }

    let settlement = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "DURATION settlement is not numeric".to_string(),
            ));
        },
    };
    let maturity = match number_arg(ctx, current_sheet, &args[1])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "DURATION maturity is not numeric".to_string(),
            ));
        },
    };
    if maturity <= settlement {
        return Ok(CellValue::Error(
            "DURATION requires maturity to be after settlement".to_string(),
        ));
    }

    let coupon_rate = match number_arg(ctx, current_sheet, &args[2])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "DURATION coupon is not numeric".to_string(),
            ));
        },
    };
    let yld = match number_arg(ctx, current_sheet, &args[3])? {
        Some(v) => v,
        None => return Ok(CellValue::Error("DURATION yld is not numeric".to_string())),
    };
    let freq_val = match number_arg(ctx, current_sheet, &args[4])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "DURATION frequency is not numeric".to_string(),
            ));
        },
    };

    let freq = freq_val.trunc() as i32;
    if !matches!(freq, 1 | 2 | 4) {
        return Ok(CellValue::Error(
            "DURATION frequency must be 1 (annual), 2 (semiannual), or 4 (quarterly)".to_string(),
        ));
    }

    let basis = if args.len() == 6 {
        match number_arg(ctx, current_sheet, &args[5])? {
            Some(v) => v,
            None => {
                return Ok(CellValue::Error(
                    "DURATION basis is not numeric".to_string(),
                ));
            },
        }
    } else {
        0.0
    };

    let basis_int = basis.trunc() as i32;
    if !(0..=4).contains(&basis_int) {
        return Ok(CellValue::Error(
            "DURATION basis must be between 0 and 4".to_string(),
        ));
    }

    let days = maturity - settlement;
    if days <= 0.0 {
        return Ok(CellValue::Error(
            "DURATION requires maturity to be after settlement".to_string(),
        ));
    }

    let year_days = match basis_int {
        0 | 2 => 360.0,
        _ => 365.0,
    };

    let years = days / year_days;
    let mut nper = (years * freq as f64).round() as i32;
    if nper <= 0 {
        nper = 1;
    }

    // Coupon bond Macaulay duration with normalized par value of 100.
    let coupon = coupon_rate * 100.0 / freq as f64;
    let per_period_yield = yld / freq as f64;

    let n = nper as i32;
    let mut pv_total = 0.0f64;
    let mut weighted_sum = 0.0f64;

    for k in 1..=n {
        let cf = if k < n { coupon } else { coupon + 100.0 };
        let t_years = k as f64 / freq as f64;
        let discount = (1.0 + per_period_yield).powi(k);
        let pv = cf / discount;
        pv_total += pv;
        weighted_sum += t_years * pv;
    }

    if pv_total == 0.0 {
        return Ok(CellValue::Error(
            "DURATION present value is zero".to_string(),
        ));
    }

    let duration = weighted_sum / pv_total;
    Ok(CellValue::Float(duration))
}

fn number_arg<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    expr: &Expr,
) -> Result<Option<f64>> {
    let v = evaluate_expression(ctx, current_sheet, expr)?;
    Ok(to_number(&v))
}

fn present_value(rate: f64, nper: f64, pmt: f64, fv: f64, typ: f64) -> f64 {
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

fn future_value(rate: f64, nper: f64, pmt: f64, pv: f64, typ: f64) -> f64 {
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

fn solve_rate(nper: f64, pmt: f64, pv: f64, fv: f64, typ: f64, guess: f64) -> Option<f64> {
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

fn numerical_derivative<F>(f: F, x: f64) -> f64
where
    F: Fn(f64) -> f64,
{
    let h = 1e-5;
    let f1 = f(x + h);
    let f2 = f(x - h);
    (f1 - f2) / (2.0 * h)
}

fn solve_irr(cash_flows: &[f64], guess: f64) -> Option<f64> {
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

fn npv_for_irr(rate: f64, cash_flows: &[f64]) -> f64 {
    let mut total = 0.0;
    for (i, cf) in cash_flows.iter().enumerate() {
        let t = i as f64;
        let denom = (1.0 + rate).powf(t);
        total += cf / denom;
    }
    total
}

fn xnpv(rate: f64, cash_flows: &[f64], dates: &[f64]) -> f64 {
    let base_date = dates[0];
    let mut total = 0.0;
    for (cf, d) in cash_flows.iter().zip(dates.iter()) {
        let t = (d - base_date) / 365.0;
        let denom = (1.0 + rate).powf(t);
        total += cf / denom;
    }
    total
}

fn solve_xirr(cash_flows: &[f64], dates: &[f64], guess: f64) -> Option<f64> {
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
