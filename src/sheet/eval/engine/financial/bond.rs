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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::parser::Expr;

    fn num_expr(n: f64) -> Expr {
        if n == n.floor() {
            Expr::Literal(CellValue::Int(n as i64))
        } else {
            Expr::Literal(CellValue::Float(n))
        }
    }

    #[tokio::test]
    async fn test_eval_yield_basic() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // YIELD with settlement=0, maturity=365, rate=5%, pr=95, redemption=100, frequency=2
        let args = vec![
            num_expr(0.0),
            num_expr(365.0),
            num_expr(0.05),
            num_expr(95.0),
            num_expr(100.0),
            num_expr(2.0),
        ];
        let result = eval_yield(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.0 && v < 0.15),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_yield_with_basis() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // YIELD with basis=1 (actual/actual)
        let args = vec![
            num_expr(0.0),
            num_expr(365.0),
            num_expr(0.05),
            num_expr(95.0),
            num_expr(100.0),
            num_expr(2.0),
            num_expr(1.0),
        ];
        let result = eval_yield(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.0 && v < 0.15),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_yield_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0), num_expr(365.0)];
        let result = eval_yield(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 6 or 7")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_yield_invalid_frequency() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(0.0),
            num_expr(365.0),
            num_expr(0.05),
            num_expr(95.0),
            num_expr(100.0),
            num_expr(3.0), // Invalid: must be 1, 2, or 4
        ];
        let result = eval_yield(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("frequency must be")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_yield_maturity_before_settlement() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(365.0),
            num_expr(0.0), // Maturity before settlement
            num_expr(0.05),
            num_expr(95.0),
            num_expr(100.0),
            num_expr(2.0),
        ];
        let result = eval_yield(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("maturity to be after settlement")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_duration_basic() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // DURATION with settlement=0, maturity=730 (2 years), coupon=5%, yield=6%, frequency=2
        let args = vec![
            num_expr(0.0),
            num_expr(730.0),
            num_expr(0.05),
            num_expr(0.06),
            num_expr(2.0),
        ];
        let result = eval_duration(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // Duration should be around 1.9 years for a 2-year bond
            CellValue::Float(v) => assert!(v > 1.0 && v < 3.0),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_duration_with_basis() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(0.0),
            num_expr(730.0),
            num_expr(0.05),
            num_expr(0.06),
            num_expr(2.0),
            num_expr(1.0), // Actual/actual basis
        ];
        let result = eval_duration(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 1.0 && v < 3.0),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_duration_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0), num_expr(365.0)];
        let result = eval_duration(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 5 or 6")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_duration_invalid_frequency() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(0.0),
            num_expr(730.0),
            num_expr(0.05),
            num_expr(0.06),
            num_expr(3.0), // Invalid frequency
        ];
        let result = eval_duration(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("frequency error")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_accrint_basic() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // ACCRINT: issue=0, first_interest=180, settlement=90, rate=5%, par=1000
        let args = vec![
            num_expr(0.0),
            num_expr(180.0),
            num_expr(90.0),
            num_expr(0.05),
            num_expr(1000.0),
        ];
        let result = eval_accrint(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // Accrued interest for 90 days at 5% on $1000 par
            CellValue::Float(v) => {
                // Expected: (1000 * 0.05 * 90) / 360 = 12.5
                assert!((v - 12.5).abs() < 1.0)
            },
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_accrint_with_basis() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // ACCRINT with basis=1 (actual/365)
        let args = vec![
            num_expr(0.0),
            num_expr(180.0),
            num_expr(90.0),
            num_expr(0.05),
            num_expr(1000.0),
            num_expr(1.0), // Basis
        ];
        let result = eval_accrint(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => {
                // With 365 day year: (1000 * 0.05 * 90) / 365 ≈ 12.33
                assert!((v - 12.33).abs() < 1.0)
            },
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_accrint_invalid_rate() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(0.0),
            num_expr(180.0),
            num_expr(90.0),
            num_expr(-0.05), // Negative rate
            num_expr(1000.0),
        ];
        let result = eval_accrint(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_accrint_settlement_before_issue() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(90.0),
            num_expr(180.0),
            num_expr(0.0), // Settlement before issue
            num_expr(0.05),
            num_expr(1000.0),
        ];
        let result = eval_accrint(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_accrintm_basic() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // ACCRINTM: issue=0, settlement=180, rate=5%, par=1000
        let args = vec![
            num_expr(0.0),
            num_expr(180.0),
            num_expr(0.05),
            num_expr(1000.0),
        ];
        let result = eval_accrintm(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // Accrued interest for 180 days at 5% on $1000 par
            CellValue::Float(v) => {
                // Expected: (1000 * 0.05 * 180) / 360 = 25.0
                assert!((v - 25.0).abs() < 1.0)
            },
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_accrintm_default_par() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // ACCRINTM with default par=1000
        let args = vec![num_expr(0.0), num_expr(180.0), num_expr(0.05)];
        let result = eval_accrintm(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.0),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_yielddisc_basic() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // YIELDDISC: settlement=0, maturity=180, pr=95, redemption=100
        let args = vec![
            num_expr(0.0),
            num_expr(180.0),
            num_expr(95.0),
            num_expr(100.0),
        ];
        let result = eval_yielddisc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // Yield for discounted security
            CellValue::Float(v) => assert!(v > 0.0 && v < 0.5),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_yielddisc_invalid_price() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(0.0),
            num_expr(180.0),
            num_expr(-95.0), // Negative price
            num_expr(100.0),
        ];
        let result = eval_yielddisc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_yielddisc_settlement_after_maturity() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(180.0),
            num_expr(0.0), // Maturity before settlement
            num_expr(95.0),
            num_expr(100.0),
        ];
        let result = eval_yielddisc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_yieldmat_basic() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // YIELDMAT: settlement=90, maturity=365, issue=0, rate=5%, pr=98
        let args = vec![
            num_expr(90.0),
            num_expr(365.0),
            num_expr(0.0),
            num_expr(0.05),
            num_expr(98.0),
        ];
        let result = eval_yieldmat(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.0 && v < 0.5),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_yieldmat_invalid_dates() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Settlement after maturity
        let args = vec![
            num_expr(365.0),
            num_expr(90.0),
            num_expr(0.0),
            num_expr(0.05),
            num_expr(98.0),
        ];
        let result = eval_yieldmat(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_disc_basic() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // DISC: settlement=0, maturity=180, pr=95, redemption=100
        let args = vec![
            num_expr(0.0),
            num_expr(180.0),
            num_expr(95.0),
            num_expr(100.0),
        ];
        let result = eval_disc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // Discount rate
            CellValue::Float(v) => assert!(v > 0.0 && v < 0.5),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_disc_invalid() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(0.0),
            num_expr(180.0),
            num_expr(-95.0), // Negative price
            num_expr(100.0),
        ];
        let result = eval_disc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_intrate_basic() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // INTRATE: settlement=0, maturity=180, investment=95000, redemption=100000
        let args = vec![
            num_expr(0.0),
            num_expr(180.0),
            num_expr(95000.0),
            num_expr(100000.0),
        ];
        let result = eval_intrate(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // Interest rate for fully invested security
            CellValue::Float(v) => assert!(v > 0.0 && v < 0.5),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_intrate_invalid() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(0.0),
            num_expr(180.0),
            num_expr(-95000.0), // Negative investment
            num_expr(100000.0),
        ];
        let result = eval_intrate(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_coupdaybs_stub() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_coupdaybs(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float(0.0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_coupdays_stub() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_coupdays(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float(0.0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_coupdaysnc_stub() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_coupdaysnc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float(0.0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_coupncd_stub() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_coupncd(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float(0.0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_coupnum_stub() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_coupnum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 0),
            _ => panic!("Expected Int(0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_couppcd_stub() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_couppcd(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float(0.0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_amordegrc_stub() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_amordegrc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float(0.0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_amorlinc_stub() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_amorlinc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float(0.0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_pricedisc_stub() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_pricedisc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float(0.0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_pricemat_stub() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_pricemat(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float(0.0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_received_stub() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_received(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float(0.0)"),
        }
    }
}
