use crate::sheet::CellValue;
use crate::sheet::Result;
use crate::sheet::eval::engine::{
    EvalCtx, evaluate_expression, flatten_range_expr, for_each_value_in_expr, to_bool, to_number,
};
use crate::sheet::eval::parser::Expr;

use super::helpers::{
    future_value, number_arg, present_value, solve_irr, solve_rate, solve_xirr, xnpv,
};

pub(crate) async fn eval_ddb(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 4 || args.len() > 5 {
        return Ok(CellValue::Error(
            "DDB expects 4 or 5 arguments (cost, salvage, life, period, [factor])".to_string(),
        ));
    }

    let cost = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("DDB cost is not numeric".to_string())),
    };
    let salvage = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("DDB salvage is not numeric".to_string())),
    };
    let life = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("DDB life is not numeric".to_string())),
    };
    let period = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("DDB period is not numeric".to_string())),
    };
    let factor = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("DDB factor is not numeric".to_string())),
        }
    } else {
        2.0
    };

    if cost <= 0.0 {
        return Ok(CellValue::Error(
            "DDB cost must be greater than 0".to_string(),
        ));
    }
    if salvage < 0.0 {
        return Ok(CellValue::Error("DDB salvage must be >= 0".to_string()));
    }
    if salvage > cost {
        return Ok(CellValue::Error("DDB salvage must be <= cost".to_string()));
    }
    if life <= 0.0 {
        return Ok(CellValue::Error(
            "DDB life must be greater than 0".to_string(),
        ));
    }
    if period <= 0.0 || period > life {
        return Ok(CellValue::Error(
            "DDB period must be > 0 and <= life".to_string(),
        ));
    }
    if factor <= 0.0 {
        return Ok(CellValue::Error(
            "DDB factor must be greater than 0".to_string(),
        ));
    }

    fn depreciation_step(book_value: f64, factor: f64, life: f64, salvage: f64) -> f64 {
        let mut depreciation = book_value * factor / life;
        let max_allowed = book_value - salvage;
        if depreciation > max_allowed {
            depreciation = max_allowed;
        }
        depreciation
    }

    let mut book_value = cost;
    let full_periods = period.floor();
    let fractional = period - full_periods;
    let mut depreciation = 0.0;

    if full_periods > 0.0 {
        for _ in 0..full_periods as i64 {
            depreciation = depreciation_step(book_value, factor, life, salvage);
            book_value -= depreciation;
        }
    }

    if fractional > 0.0 {
        let mut partial = depreciation_step(book_value, factor, life, salvage);
        partial *= fractional;
        let max_allowed = book_value - salvage;
        if partial > max_allowed {
            partial = max_allowed;
        }
        depreciation = partial;
    } else if full_periods == 0.0 {
        // Period less than 1: compute proportional first period only.
        let mut partial = depreciation_step(book_value, factor, life, salvage);
        partial *= period;
        let max_allowed = book_value - salvage;
        if partial > max_allowed {
            partial = max_allowed;
        }
        depreciation = partial;
    }

    Ok(CellValue::Float(depreciation))
}

pub(crate) async fn eval_pv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 5 {
        return Ok(CellValue::Error(
            "PV expects 3 to 5 arguments (rate, nper, pmt, [fv], [type])".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("PV rate is not numeric".to_string())),
    };
    let nper = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("PV nper is not numeric".to_string())),
    };
    let pmt = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("PV pmt is not numeric".to_string())),
    };

    let fv = if args.len() >= 4 {
        match number_arg(ctx, current_sheet, &args[3]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("PV fv is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let typ = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("PV type is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let pv = present_value(rate, nper, pmt, fv, typ);
    Ok(CellValue::Float(pv))
}

pub(crate) async fn eval_fv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 5 {
        return Ok(CellValue::Error(
            "FV expects 3 to 5 arguments (rate, nper, pmt, [pv], [type])".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("FV rate is not numeric".to_string())),
    };
    let nper = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("FV nper is not numeric".to_string())),
    };
    let pmt = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("FV pmt is not numeric".to_string())),
    };

    let pv = if args.len() >= 4 {
        match number_arg(ctx, current_sheet, &args[3]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("FV pv is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let typ = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("FV type is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let fv = future_value(rate, nper, pmt, pv, typ);
    Ok(CellValue::Float(fv))
}

pub(crate) async fn eval_rate(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 6 {
        return Ok(CellValue::Error(
            "RATE expects 3 to 6 arguments (nper, pmt, pv, [fv], [type], [guess])".to_string(),
        ));
    }

    let nper = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("RATE nper is not numeric".to_string())),
    };
    let pmt = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("RATE pmt is not numeric".to_string())),
    };
    let pv = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("RATE pv is not numeric".to_string())),
    };

    let fv = if args.len() >= 4 {
        match number_arg(ctx, current_sheet, &args[3]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("RATE fv is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let typ = if args.len() >= 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("RATE type is not numeric".to_string())),
        }
    } else {
        0.0
    };

    let guess = if args.len() == 6 {
        match number_arg(ctx, current_sheet, &args[5]).await? {
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

pub(crate) async fn eval_pduration(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "PDURATION expects 3 arguments (rate, pv, fv)".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let pv = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let fv = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    if rate <= 0.0 || pv <= 0.0 || fv <= 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let ratio = fv / pv;
    if ratio <= 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let denom = (1.0 + rate).ln();
    if denom == 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let periods = ratio.ln() / denom;
    if !periods.is_finite() {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    Ok(CellValue::Float(periods))
}

pub(crate) async fn eval_npv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 {
        return Ok(CellValue::Error(
            "NPV expects at least 2 arguments (rate, value1, [value2], ...)".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("NPV rate is not numeric".to_string())),
    };

    let mut total = 0.0f64;
    let mut period = 1.0f64;

    for expr in &args[1..] {
        let range = flatten_range_expr(ctx, current_sheet, expr).await?;
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

pub(crate) async fn eval_irr(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "IRR expects 1 or 2 arguments (values, [guess])".to_string(),
        ));
    }

    let range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
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
        match number_arg(ctx, current_sheet, &args[1]).await? {
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

pub(crate) async fn eval_rri(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "RRI expects 3 arguments (nper, pv, fv)".to_string(),
        ));
    }

    let nper = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let pv = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let fv = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    if nper <= 0.0 || pv <= 0.0 || fv <= 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let growth = fv / pv;
    if growth <= 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let rate = growth.powf(1.0 / nper) - 1.0;
    if !rate.is_finite() {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    Ok(CellValue::Float(rate))
}

pub(crate) async fn eval_xnpv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "XNPV expects 3 arguments (rate, values, dates)".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("XNPV rate is not numeric".to_string())),
    };

    let values_range = flatten_range_expr(ctx, current_sheet, &args[1]).await?;
    let dates_range = flatten_range_expr(ctx, current_sheet, &args[2]).await?;

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

pub(crate) async fn eval_sln(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "SLN expects 3 arguments (cost, salvage, life)".to_string(),
        ));
    }
    let cost = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let salvage = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let life = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    Ok(CellValue::Float((cost - salvage) / life))
}

pub(crate) async fn eval_syd(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 4 {
        return Ok(CellValue::Error(
            "SYD expects 4 arguments (cost, salvage, life, per)".to_string(),
        ));
    }
    let cost = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let salvage = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let life = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let per = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) if v >= 1.0 && v <= life => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let syd = ((cost - salvage) * (life - per + 1.0) * 2.0) / (life * (life + 1.0));
    Ok(CellValue::Float(syd))
}

pub(crate) async fn eval_db(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 4 || args.len() > 5 {
        return Ok(CellValue::Error(
            "DB expects 4 or 5 arguments (cost, salvage, life, period, [month])".to_string(),
        ));
    }
    let cost = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let salvage = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let life = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let period = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) if v >= 1.0 && v <= life + 1.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let month = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) if (1.0..=12.0).contains(&v) => v,
            Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
            None => 12.0,
        }
    } else {
        12.0
    };

    if salvage < 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let rate = ((1.0 - (salvage / cost).powf(1.0 / life)) * 1000.0).round() / 1000.0;

    let mut book_value = cost;
    let mut depreciation = 0.0;

    for p in 1..=(period.trunc() as i32) {
        if p == 1 {
            depreciation = cost * rate * month / 12.0;
        } else {
            depreciation = book_value * rate;
        }
        book_value -= depreciation;
    }

    if period.trunc() < period {
        // Handle fractional period if necessary, though Excel DB usually works on discrete periods
    }

    if period > life && month < 12.0 {
        // Last partial year if applicable
        let last_year_depr = book_value * rate * (12.0 - month) / 12.0;
        if period as i32 == (life as i32 + 1) {
            depreciation = last_year_depr;
        }
    }

    Ok(CellValue::Float(depreciation))
}

pub(crate) async fn eval_nominal(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "NOMINAL expects 2 arguments (effect_rate, npery)".to_string(),
        ));
    }
    let effect_rate = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let npery = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) if v >= 1.0 => v.trunc(),
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let nominal = npery * ((1.0 + effect_rate).powf(1.0 / npery) - 1.0);
    Ok(CellValue::Float(nominal))
}

pub(crate) async fn eval_effect(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "EFFECT expects 2 arguments (nominal_rate, npery)".to_string(),
        ));
    }
    let nominal_rate = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let npery = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) if v >= 1.0 => v.trunc(),
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let effect = (1.0 + nominal_rate / npery).powf(npery) - 1.0;
    Ok(CellValue::Float(effect))
}

pub(crate) async fn eval_xirr(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "XIRR expects 2 or 3 arguments (values, dates, [guess])".to_string(),
        ));
    }

    let values_range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let dates_range = flatten_range_expr(ctx, current_sheet, &args[1]).await?;

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
        match number_arg(ctx, current_sheet, &args[2]).await? {
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

pub(crate) async fn eval_pmt(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 5 {
        return Ok(CellValue::Error(
            "PMT expects 3 to 5 arguments (rate, nper, pv, [fv], [type])".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let nper = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let pv = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let fv = if args.len() >= 4 {
        match number_arg(ctx, current_sheet, &args[3]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        0.0
    };
    let typ = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        0.0
    };

    let pmt = if rate == 0.0 {
        -(pv + fv) / nper
    } else {
        let factor = (1.0 + rate).powf(nper);
        -(pv * factor + fv) * rate / (factor - 1.0) / (1.0 + rate * typ)
    };

    Ok(CellValue::Float(pmt))
}

pub(crate) async fn eval_ipmt(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 4 || args.len() > 6 {
        return Ok(CellValue::Error(
            "IPMT expects 4 to 6 arguments (rate, per, nper, pv, [fv], [type])".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let per = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let nper = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let pv = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let fv = if args.len() >= 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        0.0
    };
    let typ = if args.len() == 6 {
        match number_arg(ctx, current_sheet, &args[5]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        0.0
    };

    if per < 1.0 || per > nper {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let pmt_args = [
        Expr::Literal(CellValue::Float(rate)),
        Expr::Literal(CellValue::Float(nper)),
        Expr::Literal(CellValue::Float(pv)),
        Expr::Literal(CellValue::Float(fv)),
        Expr::Literal(CellValue::Float(typ)),
    ];
    let pmt = match eval_pmt(ctx, current_sheet, &pmt_args).await? {
        CellValue::Float(v) => v,
        _ => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let interest = if per == 1.0 && typ == 1.0 {
        0.0
    } else {
        let fv_prev = future_value(rate, per - 1.0, pmt, pv, typ);
        fv_prev * rate
    };

    Ok(CellValue::Float(interest))
}

pub(crate) async fn eval_ppmt(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 4 || args.len() > 6 {
        return Ok(CellValue::Error(
            "PPMT expects 4 to 6 arguments (rate, per, nper, pv, [fv], [type])".to_string(),
        ));
    }

    let pmt = match eval_pmt(ctx, current_sheet, args).await? {
        CellValue::Float(v) => v,
        other => return Ok(other),
    };

    let ipmt_args = args; // IPMT has same args as PPMT
    let ipmt = match eval_ipmt(ctx, current_sheet, ipmt_args).await? {
        CellValue::Float(v) => v,
        other => return Ok(other),
    };

    Ok(CellValue::Float(pmt - ipmt))
}

pub(crate) async fn eval_ispmt(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 4 {
        return Ok(CellValue::Error(
            "ISPMT expects 4 arguments (rate, per, nper, pv)".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let per = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let nper = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let pv = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    // ISPMT calculates interest for a loan with even principal payments
    let interest = -pv * rate * (1.0 - per / nper);
    Ok(CellValue::Float(interest))
}

pub(crate) async fn eval_nper(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 5 {
        return Ok(CellValue::Error(
            "NPER expects 3 to 5 arguments (rate, pmt, pv, [fv], [type])".to_string(),
        ));
    }

    let rate = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let pmt = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let pv = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let fv = if args.len() >= 4 {
        match number_arg(ctx, current_sheet, &args[3]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        0.0
    };
    let typ = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        0.0
    };

    if rate == 0.0 {
        if pmt == 0.0 {
            return Ok(CellValue::Error("#DIV/0!".to_string()));
        }
        return Ok(CellValue::Float(-(pv + fv) / pmt));
    }

    let num = pmt * (1.0 + rate * typ) - fv * rate;
    let den = pmt * (1.0 + rate * typ) + pv * rate;

    if num <= 0.0 || den <= 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let nper = (num / den).ln() / (1.0 + rate).ln();
    Ok(CellValue::Float(nper))
}

pub(crate) async fn eval_mirr(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "MIRR expects 3 arguments (values, finance_rate, reinvest_rate)".to_string(),
        ));
    }

    let range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let mut cash_flows = Vec::new();
    for v in range.values {
        if let Some(n) = to_number(&v) {
            cash_flows.push(n);
        }
    }

    if cash_flows.len() < 2 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let has_pos = cash_flows.iter().any(|v| *v > 0.0);
    let has_neg = cash_flows.iter().any(|v| *v < 0.0);
    if !has_pos || !has_neg {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let finance_rate = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let reinvest_rate = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let mut npv_pos = 0.0;
    let mut npv_neg = 0.0;
    for (i, &cf) in cash_flows.iter().enumerate() {
        if cf >= 0.0 {
            npv_pos += cf / (1.0 + reinvest_rate).powi(i as i32);
        } else {
            npv_neg += cf / (1.0 + finance_rate).powi(i as i32);
        }
    }

    if npv_neg == 0.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let n = cash_flows.len() as f64;
    let mirr = ((-npv_pos * (1.0 + reinvest_rate).powf(n - 1.0))
        / (npv_neg * (1.0 + finance_rate)))
        .powf(1.0 / (n - 1.0))
        - 1.0;

    Ok(CellValue::Float(mirr))
}

pub(crate) async fn eval_fvschedule(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "FVSCHEDULE expects 2 arguments (principal, schedule)".to_string(),
        ));
    }

    let principal = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let mut result = principal;
    for_each_value_in_expr(ctx, current_sheet, &args[1], |val| {
        if let Some(rate) = to_number(val) {
            result *= 1.0 + rate;
        }
        Ok(())
    })
    .await?;

    Ok(CellValue::Float(result))
}

pub(crate) async fn eval_dollarde(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("DOLLARDE expects 2 arguments".to_string()));
    }
    let dollar = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let fraction = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v.trunc() as i32,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    if fraction < 0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    if fraction == 0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let int_part = dollar.trunc();
    let rem = (dollar - int_part).abs();

    let digits = (fraction as f64).log10().ceil() as i32;
    let result = int_part + (rem * 10.0f64.powi(digits)) / fraction as f64;

    Ok(CellValue::Float(result))
}

pub(crate) async fn eval_dollarfr(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("DOLLARFR expects 2 arguments".to_string()));
    }
    let dollar = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let fraction = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v.trunc() as i32,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    if fraction < 0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    if fraction == 0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let int_part = dollar.trunc();
    let rem = (dollar - int_part).abs();
    let digits = (fraction as f64).log10().ceil() as i32;
    let result = int_part + (rem * fraction as f64) / 10.0f64.powi(digits);
    Ok(CellValue::Float(result))
}

pub(crate) async fn eval_vdb(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 5 || args.len() > 7 {
        return Ok(CellValue::Error(
            "VDB expects 5 to 7 arguments (cost, salvage, life, start_period, end_date, [factor], [no_switch])"
                .to_string(),
        ));
    }

    let cost = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let salvage = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let life = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let start_period = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let end_period = match number_arg(ctx, current_sheet, &args[4]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let factor = if args.len() >= 6 {
        (number_arg(ctx, current_sheet, &args[5]).await?).unwrap_or(2.0)
    } else {
        2.0
    };

    let no_switch = if args.len() == 7 {
        to_bool(&evaluate_expression(ctx, current_sheet, &args[6]).await?)
    } else {
        false
    };

    if cost < 0.0
        || salvage < 0.0
        || life <= 0.0
        || start_period < 0.0
        || end_period < start_period
        || end_period > life
        || factor <= 0.0
    {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    // Simplified VDB implementation: variable declining balance
    // This is a complex function involving potential switch to SLN.
    // For now, implementing a basic version.
    let rate = factor / life;
    let mut current_value = cost;
    let mut total_depreciation = 0.0;

    for i in 0..(end_period.ceil() as i32) {
        let period_start = i as f64;
        let period_end = (i + 1) as f64;

        let p_start = start_period.max(period_start);
        let p_end = end_period.min(period_end);

        if p_start < p_end {
            let period_depr = if !no_switch
                && (current_value - salvage) < (current_value / (life - period_start))
            {
                // Switch to SLN
                (current_value - salvage) / (life - period_start)
            } else {
                current_value * rate
            };

            let fraction = p_end - p_start;
            let amount = (period_depr * fraction).min(current_value - salvage);
            total_depreciation += amount;
        }

        let full_period_depr =
            if !no_switch && (current_value - salvage) < (current_value / (life - period_start)) {
                (current_value - salvage) / (life - period_start)
            } else {
                current_value * rate
            };
        current_value -= full_period_depr.min(current_value - salvage);
    }

    Ok(CellValue::Float(total_depreciation))
}
