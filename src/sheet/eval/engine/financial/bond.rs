use crate::sheet::eval::engine::EvalCtx;
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::helpers::{number_arg, solve_irr};

pub(crate) async fn eval_yield(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 6 || args.len() > 7 {
        return Ok(CellValue::Error(
            "YIELD expects 6 or 7 arguments (settlement, maturity, rate, pr, redemption, frequency, [basis])"
                .to_string(),
        ));
    }

    let settlement = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "YIELD settlement is not numeric".to_string(),
            ));
        },
    };
    let maturity = match number_arg(ctx, current_sheet, &args[1]).await? {
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

    let rate = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("YIELD rate is not numeric".to_string())),
    };
    let price = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("YIELD pr is not numeric".to_string())),
    };
    let redemption = match number_arg(ctx, current_sheet, &args[4]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "YIELD redemption is not numeric".to_string(),
            ));
        },
    };
    let freq_val = match number_arg(ctx, current_sheet, &args[5]).await? {
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
            "YIELD frequency must be 1, 2, or 4".to_string(),
        ));
    }

    let basis = if args.len() == 7 {
        match number_arg(ctx, current_sheet, &args[6]).await? {
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
    let year_days = match basis_int {
        0 | 2 => 360.0,
        _ => 365.0,
    };

    let years = days / year_days;
    let mut nper = (years * freq as f64).round() as i32;
    if nper <= 0 {
        nper = 1;
    }

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
        None => return Ok(CellValue::Error("YIELD failed to converge".to_string())),
    };

    Ok(CellValue::Float(per_period_yield * freq as f64))
}

pub(crate) async fn eval_duration(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 5 || args.len() > 6 {
        return Ok(CellValue::Error(
            "DURATION expects 5 or 6 arguments".to_string(),
        ));
    }

    let settlement = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("DURATION settlement error".to_string())),
    };
    let maturity = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("DURATION maturity error".to_string())),
    };
    if maturity <= settlement {
        return Ok(CellValue::Error("DURATION date error".to_string()));
    }

    let coupon_rate = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("DURATION coupon error".to_string())),
    };
    let yld = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("DURATION yld error".to_string())),
    };
    let freq_val = match number_arg(ctx, current_sheet, &args[4]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("DURATION frequency error".to_string())),
    };

    let freq = freq_val.trunc() as i32;
    if !matches!(freq, 1 | 2 | 4) {
        return Ok(CellValue::Error("DURATION frequency error".to_string()));
    }

    let basis = if args.len() == 6 {
        match number_arg(ctx, current_sheet, &args[5]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("DURATION basis error".to_string())),
        }
    } else {
        0.0
    };

    let year_days = match basis.trunc() as i32 {
        0 | 2 => 360.0,
        _ => 365.0,
    };

    let years = (maturity - settlement) / year_days;
    let nper = (years * freq as f64).round() as i32;
    let n = nper.max(1);

    let coupon = coupon_rate * 100.0 / freq as f64;
    let per_period_yield = yld / freq as f64;

    let mut pv_total = 0.0f64;
    let mut weighted_sum = 0.0f64;

    for k in 1..=n {
        let cf = if k < n { coupon } else { coupon + 100.0 };
        let t_years = k as f64 / freq as f64;
        let pv = cf / (1.0 + per_period_yield).powi(k);
        pv_total += pv;
        weighted_sum += t_years * pv;
    }

    if pv_total == 0.0 {
        return Ok(CellValue::Error("DURATION error".to_string()));
    }
    Ok(CellValue::Float(weighted_sum / pv_total))
}

pub(crate) async fn eval_accrint(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 4 || args.len() > 7 {
        return Ok(CellValue::Error(
            "ACCRINT expects 4 to 7 arguments".to_string(),
        ));
    }
    let issue = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let settlement = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let rate = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let par = if args.len() >= 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        1000.0
    };

    if rate <= 0.0 || par <= 0.0 || settlement <= issue {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let basis = if args.len() == 7 {
        match number_arg(ctx, current_sheet, &args[6]).await? {
            Some(v) => v.trunc() as i32,
            None => 0,
        }
    } else {
        0
    };
    let days_in_year = match basis {
        0 | 2 | 4 => 360.0,
        1 | 3 => 365.0,
        _ => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    Ok(CellValue::Float(
        (par * rate * (settlement - issue)) / days_in_year,
    ))
}

pub(crate) async fn eval_accrintm(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 5 {
        return Ok(CellValue::Error(
            "ACCRINTM expects 3 to 5 arguments".to_string(),
        ));
    }
    let issue = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let settlement = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let rate = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let par = if args.len() >= 4 {
        match number_arg(ctx, current_sheet, &args[3]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        1000.0
    };

    if rate <= 0.0 || par <= 0.0 || settlement <= issue {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let basis = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v.trunc() as i32,
            None => 0,
        }
    } else {
        0
    };
    let days_in_year = match basis {
        0 | 2 | 4 => 360.0,
        1 | 3 => 365.0,
        _ => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    Ok(CellValue::Float(
        (par * rate * (settlement - issue)) / days_in_year,
    ))
}

pub(crate) async fn eval_yielddisc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 4 || args.len() > 5 {
        return Ok(CellValue::Error(
            "YIELDDISC expects 4 or 5 arguments".to_string(),
        ));
    }
    let settlement = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let maturity = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let pr = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let redemption = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    if pr <= 0.0 || redemption <= 0.0 || settlement >= maturity {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let basis = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v.trunc() as i32,
            None => 0,
        }
    } else {
        0
    };
    let days_in_year = match basis {
        0 | 2 | 4 => 360.0,
        1 | 3 => 365.0,
        _ => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let fraction = (maturity - settlement) / days_in_year;
    Ok(CellValue::Float((redemption - pr) / (pr * fraction)))
}

pub(crate) async fn eval_yieldmat(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 5 || args.len() > 6 {
        return Ok(CellValue::Error(
            "YIELDMAT expects 5 or 6 arguments".to_string(),
        ));
    }
    let settlement = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let maturity = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let issue = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let rate = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let pr = match number_arg(ctx, current_sheet, &args[4]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    if pr <= 0.0 || rate < 0.0 || settlement >= maturity || issue > settlement {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let basis = if args.len() == 6 {
        match number_arg(ctx, current_sheet, &args[5]).await? {
            Some(v) => v.trunc() as i32,
            None => 0,
        }
    } else {
        0
    };
    let days_in_year = match basis {
        0 | 2 | 4 => 360.0,
        1 | 3 => 365.0,
        _ => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let issue_to_mat = (maturity - issue) / days_in_year;
    let issue_to_settle = (settlement - issue) / days_in_year;
    let settle_to_mat = (maturity - settlement) / days_in_year;

    let redemption_val = 100.0 * (1.0 + issue_to_mat * rate);
    let price_with_accrued = pr + 100.0 * issue_to_settle * rate;

    Ok(CellValue::Float(
        (redemption_val / price_with_accrued - 1.0) / settle_to_mat,
    ))
}

pub(crate) async fn eval_coupdaybs(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    _args: &[Expr],
) -> Result<CellValue> {
    Ok(CellValue::Float(0.0))
}
pub(crate) async fn eval_coupdays(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    _args: &[Expr],
) -> Result<CellValue> {
    Ok(CellValue::Float(0.0))
}
pub(crate) async fn eval_coupdaysnc(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    _args: &[Expr],
) -> Result<CellValue> {
    Ok(CellValue::Float(0.0))
}
pub(crate) async fn eval_coupncd(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    _args: &[Expr],
) -> Result<CellValue> {
    Ok(CellValue::Float(0.0))
}
pub(crate) async fn eval_coupnum(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    _args: &[Expr],
) -> Result<CellValue> {
    Ok(CellValue::Int(0))
}
pub(crate) async fn eval_couppcd(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    _args: &[Expr],
) -> Result<CellValue> {
    Ok(CellValue::Float(0.0))
}

pub(crate) async fn eval_disc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 4 || args.len() > 5 {
        return Ok(CellValue::Error(
            "DISC expects 4 or 5 arguments".to_string(),
        ));
    }
    let settlement = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let maturity = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let pr = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let redemption = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    if pr <= 0.0 || redemption <= 0.0 || settlement >= maturity {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let basis = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v.trunc() as i32,
            None => 0,
        }
    } else {
        0
    };
    let days_in_year = match basis {
        0 | 2 | 4 => 360.0,
        1 | 3 => 365.0,
        _ => return Ok(CellValue::Error("#NUM!".to_string())),
    };
    Ok(CellValue::Float(
        (redemption - pr) / redemption * (days_in_year / (maturity - settlement)),
    ))
}

pub(crate) async fn eval_intrate(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 4 || args.len() > 5 {
        return Ok(CellValue::Error(
            "INTRATE expects 4 or 5 arguments".to_string(),
        ));
    }
    let settlement = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let maturity = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let investment = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let redemption = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    if investment <= 0.0 || redemption <= 0.0 || settlement >= maturity {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let basis = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v.trunc() as i32,
            None => 0,
        }
    } else {
        0
    };
    let days_in_year = match basis {
        0 | 2 | 4 => 360.0,
        1 | 3 => 365.0,
        _ => return Ok(CellValue::Error("#NUM!".to_string())),
    };
    Ok(CellValue::Float(
        (redemption - investment) / investment * (days_in_year / (maturity - settlement)),
    ))
}

pub(crate) async fn eval_amordegrc(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    _args: &[Expr],
) -> Result<CellValue> {
    Ok(CellValue::Float(0.0))
}
pub(crate) async fn eval_amorlinc(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    _args: &[Expr],
) -> Result<CellValue> {
    Ok(CellValue::Float(0.0))
}
pub(crate) async fn eval_pricedisc(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    _args: &[Expr],
) -> Result<CellValue> {
    Ok(CellValue::Float(0.0))
}
pub(crate) async fn eval_pricemat(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    _args: &[Expr],
) -> Result<CellValue> {
    Ok(CellValue::Float(0.0))
}
pub(crate) async fn eval_received(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    _args: &[Expr],
) -> Result<CellValue> {
    Ok(CellValue::Float(0.0))
}
