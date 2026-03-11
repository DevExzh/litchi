use super::super::{
    EvalCtx, evaluate_expression, flatten_range_expr, for_each_value_in_expr, to_bool, to_number,
};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};
use statrs::distribution::{
    Beta, Binomial, ChiSquared, Continuous, ContinuousCDF, Discrete, DiscreteCDF, Exp,
    FisherSnedecor, Gamma, Hypergeometric, LogNormal, NegativeBinomial, Normal, Poisson, StudentsT,
    Weibull,
};

use super::helpers::number_arg;

pub(crate) async fn eval_norm_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "NORM.DIST expects 3 or 4 arguments (x, mean, standard_dev, [cumulative])".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("NORM.DIST x is not numeric".to_string()));
        },
    };
    let mean = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "NORM.DIST mean is not numeric".to_string(),
            ));
        },
    };
    let std_dev = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "NORM.DIST standard_dev is not numeric".to_string(),
            ));
        },
    };

    if std_dev <= 0.0 {
        return Ok(CellValue::Error(
            "NORM.DIST standard_dev must be > 0".to_string(),
        ));
    }

    let dist = match Normal::new(mean, std_dev) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("NORM.DIST domain error".to_string()));
        },
    };

    let cumulative = if args.len() == 4 {
        let v = evaluate_expression(ctx, current_sheet, &args[3]).await?;
        to_bool(&v)
    } else {
        true
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_norm_s_inv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error(
            "NORM.S.INV expects 1 argument (probability)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "NORM.S.INV probability is not numeric".to_string(),
            ));
        },
    };

    if !(0.0..=1.0).contains(&p) {
        return Ok(CellValue::Error(
            "NORM.S.INV probability must be between 0 and 1".to_string(),
        ));
    }

    let dist = Normal::standard();
    let value = dist.inverse_cdf(p);
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_norm_s_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "NORM.S.DIST expects 1 or 2 arguments (z, [cumulative])".to_string(),
        ));
    }

    let z = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("NORM.S.DIST z is not numeric".to_string()));
        },
    };

    let cumulative = if args.len() == 2 {
        let v = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        to_bool(&v)
    } else {
        true
    };

    let dist = Normal::standard();
    let value = if cumulative { dist.cdf(z) } else { dist.pdf(z) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_beta_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 4 || args.len() > 6 {
        return Ok(CellValue::Error(
            "BETA.DIST expects 4 to 6 arguments (x, alpha, beta, cumulative, [A], [B])".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let alpha = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let beta = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let cumulative = to_bool(&evaluate_expression(ctx, current_sheet, &args[3]).await?);

    let a = if args.len() >= 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        0.0
    };

    let b = if args.len() == 6 {
        match number_arg(ctx, current_sheet, &args[5]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        1.0
    };

    if x < a || x > b || a >= b {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let dist = match Beta::new(alpha, beta) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let scaled_x = (x - a) / (b - a);
    let value = if cumulative {
        dist.cdf(scaled_x)
    } else {
        dist.pdf(scaled_x) / (b - a)
    };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_beta_inv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 5 {
        return Ok(CellValue::Error(
            "BETA.INV expects 3 to 5 arguments (probability, alpha, beta, [A], [B])".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if (0.0..=1.0).contains(&v) => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let alpha = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let beta = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let a = if args.len() >= 4 {
        match number_arg(ctx, current_sheet, &args[3]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        0.0
    };

    let b = if args.len() == 5 {
        match number_arg(ctx, current_sheet, &args[4]).await? {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        1.0
    };

    if a >= b {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let dist = match Beta::new(alpha, beta) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let scaled_inv = dist.inverse_cdf(p);
    Ok(CellValue::Float(a + scaled_inv * (b - a)))
}

pub(crate) async fn eval_devsq(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "DEVSQ requires at least one argument".to_string(),
        ));
    }

    let mut values = Vec::new();
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                values.push(n);
            }
            Ok(())
        })
        .await?;
    }

    if values.is_empty() {
        return Ok(CellValue::Error(
            "DEVSQ requires numeric values".to_string(),
        ));
    }

    let mean = values.iter().sum::<f64>() / values.len() as f64;
    let total = values.iter().map(|v| (v - mean).powi(2)).sum::<f64>();
    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_prob(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "PROB expects 3 or 4 arguments (x_range, prob_range, lower_limit, [upper_limit])"
                .to_string(),
        ));
    }

    let x_range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let p_range = flatten_range_expr(ctx, current_sheet, &args[1]).await?;

    if x_range.values.len() != p_range.values.len() || x_range.values.is_empty() {
        return Ok(CellValue::Error(
            "PROB requires x_range and prob_range of the same non-zero length".to_string(),
        ));
    }

    let len = x_range.values.len();
    let mut xs = Vec::with_capacity(len);
    let mut ps = Vec::with_capacity(len);

    for (xv, pv) in x_range.values.iter().zip(p_range.values.iter()) {
        let x = match to_number(xv) {
            Some(v) => v,
            None => {
                return Ok(CellValue::Error(
                    "PROB x_range must contain only numeric values".to_string(),
                ));
            },
        };
        let p = match to_number(pv) {
            Some(v) => v,
            None => {
                return Ok(CellValue::Error(
                    "PROB prob_range must contain only numeric values".to_string(),
                ));
            },
        };
        if !(0.0..=1.0).contains(&p) {
            return Ok(CellValue::Error(
                "PROB prob_range values must be between 0 and 1".to_string(),
            ));
        }
        xs.push(x);
        ps.push(p);
    }

    let sum_p: f64 = ps.iter().sum();
    if (sum_p - 1.0).abs() > 1e-7 {
        return Ok(CellValue::Error(
            "PROB prob_range values must sum to 1".to_string(),
        ));
    }

    let lower = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "PROB lower_limit is not numeric".to_string(),
            ));
        },
    };

    let (lo, hi) = if args.len() == 4 {
        let upper = match number_arg(ctx, current_sheet, &args[3]).await? {
            Some(v) => v,
            None => {
                return Ok(CellValue::Error(
                    "PROB upper_limit is not numeric".to_string(),
                ));
            },
        };
        if lower <= upper {
            (lower, upper)
        } else {
            (upper, lower)
        }
    } else {
        (lower, lower)
    };

    let mut prob = 0.0f64;
    for (x, p) in xs.iter().zip(ps.iter()) {
        if *x >= lo && *x <= hi {
            prob += *p;
        }
    }

    Ok(CellValue::Float(prob))
}

pub(crate) async fn eval_chisq_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "CHISQ.DIST expects 2 or 3 arguments (x, deg_freedom, [cumulative])".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("CHISQ.DIST x is not numeric".to_string()));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "CHISQ.DIST deg_freedom is not numeric".to_string(),
            ));
        },
    };

    if df <= 0.0 {
        return Ok(CellValue::Error(
            "CHISQ.DIST deg_freedom must be > 0".to_string(),
        ));
    }

    let dist = match ChiSquared::new(df) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("CHISQ.DIST domain error".to_string()));
        },
    };

    let cumulative = if args.len() == 3 {
        let v = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        to_bool(&v)
    } else {
        true
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_chisq_dist_rt(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "CHISQ.DIST.RT expects 2 arguments (x, deg_freedom)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "CHISQ.DIST.RT x is not numeric".to_string(),
            ));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "CHISQ.DIST.RT deg_freedom is not numeric".to_string(),
            ));
        },
    };

    if df <= 0.0 {
        return Ok(CellValue::Error(
            "CHISQ.DIST.RT deg_freedom must be > 0".to_string(),
        ));
    }

    let dist = match ChiSquared::new(df) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("CHISQ.DIST.RT domain error".to_string()));
        },
    };

    let value = dist.sf(x);
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_chisq_inv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "CHISQ.INV expects 2 arguments (probability, deg_freedom)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "CHISQ.INV probability is not numeric".to_string(),
            ));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "CHISQ.INV deg_freedom is not numeric".to_string(),
            ));
        },
    };

    if !(0.0..=1.0).contains(&p) || df <= 0.0 {
        return Ok(CellValue::Error(
            "CHISQ.INV probability must be between 0 and 1 and deg_freedom > 0".to_string(),
        ));
    }

    let dist = match ChiSquared::new(df) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("CHISQ.INV domain error".to_string()));
        },
    };

    let value = dist.inverse_cdf(p);
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_chisq_inv_rt(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "CHISQ.INV.RT expects 2 arguments (probability, deg_freedom)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "CHISQ.INV.RT probability is not numeric".to_string(),
            ));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "CHISQ.INV.RT deg_freedom is not numeric".to_string(),
            ));
        },
    };

    if !(0.0..=1.0).contains(&p) || df <= 0.0 {
        return Ok(CellValue::Error(
            "CHISQ.INV.RT probability must be between 0 and 1 and deg_freedom > 0".to_string(),
        ));
    }

    let dist = match ChiSquared::new(df) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("CHISQ.INV.RT domain error".to_string()));
        },
    };

    let value = dist.inverse_cdf(1.0 - p);
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_t_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "T.DIST expects 3 arguments (x, deg_freedom, cumulative)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("T.DIST x is not numeric".to_string()));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "T.DIST deg_freedom is not numeric".to_string(),
            ));
        },
    };

    if df <= 0.0 {
        return Ok(CellValue::Error(
            "T.DIST deg_freedom must be > 0".to_string(),
        ));
    }

    let cumulative = {
        let v = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        to_bool(&v)
    };

    let dist = match StudentsT::new(0.0, 1.0, df) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("T.DIST domain error".to_string()));
        },
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_t_dist_2t(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "T.DIST.2T expects 2 arguments (x, deg_freedom)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("T.DIST.2T x is not numeric".to_string()));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "T.DIST.2T deg_freedom is not numeric".to_string(),
            ));
        },
    };

    if df <= 0.0 {
        return Ok(CellValue::Error(
            "T.DIST.2T deg_freedom must be > 0".to_string(),
        ));
    }

    let dist = match StudentsT::new(0.0, 1.0, df) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("T.DIST.2T domain error".to_string()));
        },
    };

    let abs_x = x.abs();
    let upper_tail = dist.sf(abs_x);
    let value = 2.0 * upper_tail;
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_t_dist_rt(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "T.DIST.RT expects 2 arguments (x, deg_freedom)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("T.DIST.RT x is not numeric".to_string()));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "T.DIST.RT deg_freedom is not numeric".to_string(),
            ));
        },
    };

    if df <= 0.0 {
        return Ok(CellValue::Error(
            "T.DIST.RT deg_freedom must be > 0".to_string(),
        ));
    }

    let dist = match StudentsT::new(0.0, 1.0, df) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("T.DIST.RT domain error".to_string()));
        },
    };

    let value = dist.sf(x);
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_t_inv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "T.INV expects 2 arguments (probability, deg_freedom)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "T.INV probability is not numeric".to_string(),
            ));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "T.INV deg_freedom is not numeric".to_string(),
            ));
        },
    };

    if !(0.0..=1.0).contains(&p) || df <= 0.0 {
        return Ok(CellValue::Error(
            "T.INV probability must be between 0 and 1 and deg_freedom > 0".to_string(),
        ));
    }

    let dist = match StudentsT::new(0.0, 1.0, df) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("T.INV domain error".to_string()));
        },
    };

    let value = dist.inverse_cdf(p);
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_t_inv_2t(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "T.INV.2T expects 2 arguments (probability, deg_freedom)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "T.INV.2T probability is not numeric".to_string(),
            ));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "T.INV.2T deg_freedom is not numeric".to_string(),
            ));
        },
    };

    if !(0.0..=1.0).contains(&p) || df <= 0.0 {
        return Ok(CellValue::Error(
            "T.INV.2T probability must be between 0 and 1 and deg_freedom > 0".to_string(),
        ));
    }

    let dist = match StudentsT::new(0.0, 1.0, df) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("T.INV.2T domain error".to_string()));
        },
    };

    let tail_prob = 1.0 - p / 2.0;
    let value = dist.inverse_cdf(tail_prob);
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_f_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "F.DIST expects 3 or 4 arguments (x, deg_freedom1, deg_freedom2, [cumulative])"
                .to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("F.DIST x is not numeric".to_string()));
        },
    };
    let df1 = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.DIST deg_freedom1 is not numeric".to_string(),
            ));
        },
    };
    let df2 = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.DIST deg_freedom2 is not numeric".to_string(),
            ));
        },
    };

    if df1 <= 0.0 || df2 <= 0.0 {
        return Ok(CellValue::Error(
            "F.DIST deg_freedom1 and deg_freedom2 must be > 0".to_string(),
        ));
    }

    let dist = match FisherSnedecor::new(df1, df2) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("F.DIST domain error".to_string()));
        },
    };

    let cumulative = if args.len() == 4 {
        let v = evaluate_expression(ctx, current_sheet, &args[3]).await?;
        to_bool(&v)
    } else {
        true
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_f_dist_rt(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "F.DIST.RT expects 3 arguments (x, deg_freedom1, deg_freedom2)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("F.DIST.RT x is not numeric".to_string()));
        },
    };
    let df1 = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.DIST.RT deg_freedom1 is not numeric".to_string(),
            ));
        },
    };
    let df2 = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.DIST.RT deg_freedom2 is not numeric".to_string(),
            ));
        },
    };

    if df1 <= 0.0 || df2 <= 0.0 {
        return Ok(CellValue::Error(
            "F.DIST.RT deg_freedom1 and deg_freedom2 must be > 0".to_string(),
        ));
    }

    let dist = match FisherSnedecor::new(df1, df2) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("F.DIST.RT domain error".to_string()));
        },
    };

    let value = dist.sf(x);
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_f_inv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "F.INV expects 3 arguments (probability, deg_freedom1, deg_freedom2)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.INV probability is not numeric".to_string(),
            ));
        },
    };
    let df1 = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.INV deg_freedom1 is not numeric".to_string(),
            ));
        },
    };
    let df2 = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.INV deg_freedom2 is not numeric".to_string(),
            ));
        },
    };

    if !(0.0..=1.0).contains(&p) || df1 <= 0.0 || df2 <= 0.0 {
        return Ok(CellValue::Error(
            "F.INV probability must be between 0 and 1 and degrees of freedom > 0".to_string(),
        ));
    }

    let dist = match FisherSnedecor::new(df1, df2) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("F.INV domain error".to_string()));
        },
    };

    let value = dist.inverse_cdf(p);
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_f_inv_rt(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "F.INV.RT expects 3 arguments (probability, deg_freedom1, deg_freedom2)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.INV.RT probability is not numeric".to_string(),
            ));
        },
    };
    let df1 = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.INV.RT deg_freedom1 is not numeric".to_string(),
            ));
        },
    };
    let df2 = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.INV.RT deg_freedom2 is not numeric".to_string(),
            ));
        },
    };

    if !(0.0..=1.0).contains(&p) || df1 <= 0.0 || df2 <= 0.0 {
        return Ok(CellValue::Error(
            "F.INV.RT probability must be between 0 and 1 and degrees of freedom > 0".to_string(),
        ));
    }

    let dist = match FisherSnedecor::new(df1, df2) {
        Ok(d) => d,
        Err(_) => {
            return Ok(CellValue::Error("F.INV.RT domain error".to_string()));
        },
    };

    let value = dist.inverse_cdf(1.0 - p);
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_binom_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 4 {
        return Ok(CellValue::Error(
            "BINOM.DIST expects 4 arguments (number_s, trials, probability_s, cumulative)"
                .to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v.trunc() as u64,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let n = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v.trunc() as u64,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let p = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let cumulative = to_bool(&evaluate_expression(ctx, current_sheet, &args[3]).await?);

    if !(0.0..=1.0).contains(&p) || x > n {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let dist = match Binomial::new(p, n) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pmf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_hypgeom_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 5 {
        return Ok(CellValue::Error(
            "HYPGEOM.DIST expects 5 arguments (sample_s, number_sample, population_s, number_pop, cumulative)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v.trunc() as u64,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let n = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v.trunc() as u64,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let m = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v.trunc() as u64,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let population_size = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v.trunc() as u64,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let cumulative = to_bool(&evaluate_expression(ctx, current_sheet, &args[4]).await?);

    if n > population_size || m > population_size || x > n || x > m {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let dist = match Hypergeometric::new(population_size, m, n) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pmf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_negbinom_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 4 {
        return Ok(CellValue::Error(
            "NEGBINOM.DIST expects 4 arguments (number_f, number_s, probability_s, cumulative)"
                .to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v.trunc() as u64,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let r = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v.trunc() as u64,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let p = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let cumulative = to_bool(&evaluate_expression(ctx, current_sheet, &args[3]).await?);

    if !(0.0..=1.0).contains(&p) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let dist = match NegativeBinomial::new(r as f64, p) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pmf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_poisson_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "POISSON.DIST expects 3 arguments (x, mean, cumulative)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v.trunc() as u64,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let lambda = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let cumulative = to_bool(&evaluate_expression(ctx, current_sheet, &args[2]).await?);

    if lambda < 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let dist = match Poisson::new(lambda) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pmf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_confidence_norm(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "CONFIDENCE.NORM expects 3 arguments (alpha, standard_dev, size)".to_string(),
        ));
    }

    let alpha = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if v > 0.0 && v < 1.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let stdev = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let size = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v >= 1.0 => v.trunc(),
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let dist = Normal::standard();
    let z = dist.inverse_cdf(1.0 - alpha / 2.0);
    Ok(CellValue::Float(z * stdev / size.sqrt()))
}

pub(crate) async fn eval_confidence_t(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "CONFIDENCE.T expects 3 arguments (alpha, standard_dev, size)".to_string(),
        ));
    }

    let alpha = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if v > 0.0 && v < 1.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let stdev = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let size = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v >= 2.0 => v.trunc(),
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let df = size - 1.0;
    let dist = match StudentsT::new(0.0, 1.0, df) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let t = dist.inverse_cdf(1.0 - alpha / 2.0);
    Ok(CellValue::Float(t * stdev / size.sqrt()))
}

pub(crate) async fn eval_expon_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "EXPON.DIST expects 3 arguments (x, lambda, cumulative)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let lambda = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let cumulative = to_bool(&evaluate_expression(ctx, current_sheet, &args[2]).await?);

    if x < 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let dist = match Exp::new(lambda) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_gamma_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 4 {
        return Ok(CellValue::Error(
            "GAMMA.DIST expects 4 arguments (x, alpha, beta, cumulative)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let alpha = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let beta = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let cumulative = to_bool(&evaluate_expression(ctx, current_sheet, &args[3]).await?);

    if x < 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let dist = match Gamma::new(alpha, 1.0 / beta) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_gammainv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "GAMMAINV expects 3 arguments (probability, alpha, beta)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if (0.0..=1.0).contains(&v) => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let alpha = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let beta = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let dist = match Gamma::new(alpha, 1.0 / beta) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    Ok(CellValue::Float(dist.inverse_cdf(p)))
}

pub(crate) async fn eval_gammaln(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("GAMMALN expects 1 argument".to_string()));
    }
    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    Ok(CellValue::Float(statrs::function::gamma::ln_gamma(x)))
}

pub(crate) async fn eval_lognorm_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 4 {
        return Ok(CellValue::Error(
            "LOGNORM.DIST expects 4 arguments (x, mean, standard_dev, cumulative)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let mean = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let stdev = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let cumulative = to_bool(&evaluate_expression(ctx, current_sheet, &args[3]).await?);

    let dist = match LogNormal::new(mean, stdev) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_lognorm_inv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "LOGNORM.INV expects 3 arguments (probability, mean, standard_dev)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if (0.0..=1.0).contains(&v) => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let mean = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let stdev = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let dist = match LogNormal::new(mean, stdev) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    Ok(CellValue::Float(dist.inverse_cdf(p)))
}

pub(crate) async fn eval_weibull_dist(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 4 {
        return Ok(CellValue::Error(
            "WEIBULL.DIST expects 4 arguments (x, alpha, beta, cumulative)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if v >= 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let alpha = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let beta = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let cumulative = to_bool(&evaluate_expression(ctx, current_sheet, &args[3]).await?);

    let dist = match Weibull::new(alpha, beta) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_z_test(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "Z.TEST expects 2 or 3 arguments".to_string(),
        ));
    }

    let mut values = Vec::new();
    for_each_value_in_expr(ctx, current_sheet, &args[0], |v| {
        if let Some(n) = to_number(v) {
            values.push(n);
        }
        Ok(())
    })
    .await?;

    if values.is_empty() {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let x = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let n = values.len() as f64;
    let mean = values.iter().sum::<f64>() / n;

    let sigma = if args.len() == 3 {
        match number_arg(ctx, current_sheet, &args[2]).await? {
            Some(v) if v > 0.0 => v,
            Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        let sum_sq_diff: f64 = values.iter().map(|&v| (v - mean).powi(2)).sum();
        (sum_sq_diff / n).sqrt()
    };

    if sigma == 0.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let z = (mean - x) / (sigma / n.sqrt());
    let dist = Normal::standard();
    Ok(CellValue::Float(1.0 - dist.cdf(z)))
}

pub(crate) async fn eval_f_test(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("F.TEST expects 2 arguments".to_string()));
    }

    let mut values1 = Vec::new();
    for_each_value_in_expr(ctx, current_sheet, &args[0], |v| {
        if let Some(n) = to_number(v) {
            values1.push(n);
        }
        Ok(())
    })
    .await?;

    let mut values2 = Vec::new();
    for_each_value_in_expr(ctx, current_sheet, &args[1], |v| {
        if let Some(n) = to_number(v) {
            values2.push(n);
        }
        Ok(())
    })
    .await?;

    if values1.len() < 2 || values2.len() < 2 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let mean1 = values1.iter().sum::<f64>() / values1.len() as f64;
    let var1 =
        values1.iter().map(|&v| (v - mean1).powi(2)).sum::<f64>() / (values1.len() - 1) as f64;

    let mean2 = values2.iter().sum::<f64>() / values2.len() as f64;
    let var2 =
        values2.iter().map(|&v| (v - mean2).powi(2)).sum::<f64>() / (values2.len() - 1) as f64;

    if var1 == 0.0 || var2 == 0.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let (f, df1, df2) = if var1 > var2 {
        (
            var1 / var2,
            (values1.len() - 1) as f64,
            (values2.len() - 1) as f64,
        )
    } else {
        (
            var2 / var1,
            (values2.len() - 1) as f64,
            (values1.len() - 1) as f64,
        )
    };

    let dist = match FisherSnedecor::new(df1, df2) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    Ok(CellValue::Float(2.0 * (1.0 - dist.cdf(f))))
}

pub(crate) async fn eval_chisq_test(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "CHISQ.TEST expects 2 arguments (actual_range, expected_range)".to_string(),
        ));
    }

    let actual_range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let expected_range = flatten_range_expr(ctx, current_sheet, &args[1]).await?;

    if actual_range.values.len() != expected_range.values.len() || actual_range.values.is_empty() {
        return Ok(CellValue::Error("#N/A".to_string()));
    }

    let rows = actual_range.rows;
    let cols = actual_range.cols;

    if rows < 1 || cols < 1 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let mut chi_sq = 0.0;
    for (a_val, e_val) in actual_range.values.iter().zip(expected_range.values.iter()) {
        let a = to_number(a_val).unwrap_or(0.0);
        let e = to_number(e_val).unwrap_or(0.0);
        if e == 0.0 {
            return Ok(CellValue::Error("#DIV/0!".to_string()));
        }
        chi_sq += (a - e).powi(2) / e;
    }

    let df = ((rows - 1) * (cols - 1)) as f64;
    if df <= 0.0 {
        // Excel behavior for 1xN or Nx1 ranges: df = N-1
        let n = actual_range.values.len();
        let df_linear = (n - 1) as f64;
        if df_linear <= 0.0 {
            return Ok(CellValue::Error("#DIV/0!".to_string()));
        }
        let dist = match ChiSquared::new(df_linear) {
            Ok(d) => d,
            Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        };
        return Ok(CellValue::Float(dist.sf(chi_sq)));
    }

    let dist = match ChiSquared::new(df) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    Ok(CellValue::Float(dist.sf(chi_sq)))
}

pub(crate) async fn eval_t_test(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 4 {
        return Ok(CellValue::Error(
            "T.TEST expects 4 arguments (array1, array2, tails, type)".to_string(),
        ));
    }

    let array1 = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let array2 = flatten_range_expr(ctx, current_sheet, &args[1]).await?;

    let tails = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) => v.trunc() as i32,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let test_type = match number_arg(ctx, current_sheet, &args[3]).await? {
        Some(v) => v.trunc() as i32,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    if tails != 1 && tails != 2 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let mut vals1 = Vec::new();
    for v in array1.values {
        if let Some(n) = to_number(&v) {
            vals1.push(n);
        }
    }

    let mut vals2 = Vec::new();
    for v in array2.values {
        if let Some(n) = to_number(&v) {
            vals2.push(n);
        }
    }

    if vals1.is_empty() || vals2.is_empty() {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let n1 = vals1.len() as f64;
    let n2 = vals2.len() as f64;

    let p_val = match test_type {
        1 => {
            // Paired
            if vals1.len() != vals2.len() {
                return Ok(CellValue::Error("#N/A".to_string()));
            }
            let diffs: Vec<f64> = vals1.iter().zip(vals2.iter()).map(|(a, b)| a - b).collect();
            let n = diffs.len() as f64;
            let mean_diff = diffs.iter().sum::<f64>() / n;
            let sum_sq_diff: f64 = diffs.iter().map(|&d| (d - mean_diff).powi(2)).sum();
            let var_diff = sum_sq_diff / (n - 1.0);
            if var_diff == 0.0 {
                return Ok(CellValue::Error("#DIV/0!".to_string()));
            }
            let t_stat = mean_diff / (var_diff / n).sqrt();
            let df = n - 1.0;
            let dist = match StudentsT::new(0.0, 1.0, df) {
                Ok(d) => d,
                Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
            };
            let p1 = dist.sf(t_stat.abs());
            if tails == 1 { p1 } else { 2.0 * p1 }
        },
        2 => {
            // Two-sample equal variance (homoscedastic)
            let mean1 = vals1.iter().sum::<f64>() / n1;
            let var1 = vals1.iter().map(|&v| (v - mean1).powi(2)).sum::<f64>() / (n1 - 1.0);
            let mean2 = vals2.iter().sum::<f64>() / n2;
            let var2 = vals2.iter().map(|&v| (v - mean2).powi(2)).sum::<f64>() / (n2 - 1.0);

            let df = n1 + n2 - 2.0;
            let pooled_var = ((n1 - 1.0) * var1 + (n2 - 1.0) * var2) / df;
            if pooled_var == 0.0 {
                return Ok(CellValue::Error("#DIV/0!".to_string()));
            }
            let t_stat = (mean1 - mean2) / (pooled_var * (1.0 / n1 + 1.0 / n2)).sqrt();
            let dist = match StudentsT::new(0.0, 1.0, df) {
                Ok(d) => d,
                Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
            };
            let p1 = dist.sf(t_stat.abs());
            if tails == 1 { p1 } else { 2.0 * p1 }
        },
        3 => {
            // Two-sample unequal variance (heteroscedastic)
            let mean1 = vals1.iter().sum::<f64>() / n1;
            let var1 = vals1.iter().map(|&v| (v - mean1).powi(2)).sum::<f64>() / (n1 - 1.0);
            let mean2 = vals2.iter().sum::<f64>() / n2;
            let var2 = vals2.iter().map(|&v| (v - mean2).powi(2)).sum::<f64>() / (n2 - 1.0);

            let se1 = var1 / n1;
            let se2 = var2 / n2;
            if se1 + se2 == 0.0 {
                return Ok(CellValue::Error("#DIV/0!".to_string()));
            }
            let t_stat = (mean1 - mean2) / (se1 + se2).sqrt();
            let df = (se1 + se2).powi(2) / (se1.powi(2) / (n1 - 1.0) + se2.powi(2) / (n2 - 1.0));
            let dist = match StudentsT::new(0.0, 1.0, df) {
                Ok(d) => d,
                Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
            };
            let p1 = dist.sf(t_stat.abs());
            if tails == 1 { p1 } else { 2.0 * p1 }
        },
        _ => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    Ok(CellValue::Float(p_val))
}

pub(crate) async fn eval_binom_inv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "BINOM.INV expects 3 arguments (trials, probability_s, alpha)".to_string(),
        ));
    }

    let trials = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if v >= 0.0 => v.trunc() as u64,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let probability_s = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) if (0.0..=1.0).contains(&v) => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let alpha = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if (0.0..=1.0).contains(&v) => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let dist = match Binomial::new(probability_s, trials) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    // Excel's BINOM.INV is the inverse cumulative binomial distribution.
    // It returns the smallest x such that P(X <= x) >= alpha.
    // Since Binomial is discrete, we can use a simple binary search or linear search if trials is small.
    // statrs Binomial doesn't seem to have inverse_cdf for discrete.

    let mut low = 0;
    let mut high = trials;
    let mut ans = trials;

    while low <= high {
        let mid = low + (high - low) / 2;
        if dist.cdf(mid) >= alpha {
            ans = mid;
            if mid == 0 {
                break;
            }
            high = mid - 1;
        } else {
            low = mid + 1;
        }
    }

    Ok(CellValue::Int(ans as i64))
}

pub(crate) async fn eval_binom_dist_range(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "BINOM.DIST.RANGE expects 3 or 4 arguments (trials, probability_s, number_s, [number_s2])".to_string(),
        ));
    }

    let trials = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if v >= 0.0 => v.trunc() as u64,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let probability_s = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) if (0.0..=1.0).contains(&v) => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let number_s = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v >= 0.0 => v.trunc() as u64,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    if number_s > trials {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let dist = match Binomial::new(probability_s, trials) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    if args.len() == 4 {
        let number_s2 = match number_arg(ctx, current_sheet, &args[3]).await? {
            Some(v) if v >= 0.0 => v.trunc() as u64,
            Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        };
        if number_s2 < number_s || number_s2 > trials {
            return Ok(CellValue::Error("#NUM!".to_string()));
        }

        let p_up_to_s2 = dist.cdf(number_s2);
        let p_below_s1 = if number_s > 0 {
            dist.cdf(number_s - 1)
        } else {
            0.0
        };
        Ok(CellValue::Float(p_up_to_s2 - p_below_s1))
    } else {
        Ok(CellValue::Float(dist.pmf(number_s)))
    }
}

pub(crate) async fn eval_norm_inv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "NORM.INV expects 3 arguments (probability, mean, standard_dev)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) if (0.0..=1.0).contains(&v) => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let mean = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let std_dev = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let dist = match Normal::new(mean, std_dev) {
        Ok(d) => d,
        Err(_) => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    Ok(CellValue::Float(dist.inverse_cdf(p)))
}

pub(crate) async fn eval_gauss(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("GAUSS expects 1 argument".to_string()));
    }
    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let dist = Normal::standard();
    Ok(CellValue::Float(dist.cdf(x) - 0.5))
}

pub(crate) async fn eval_phi(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("PHI expects 1 argument".to_string()));
    }
    let x = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let dist = Normal::standard();
    Ok(CellValue::Float(dist.pdf(x)))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::engine::test_helpers::TestEngine;
    use crate::sheet::eval::parser::Expr;

    fn num_expr(n: f64) -> Expr {
        Expr::Literal(CellValue::Float(n))
    }

    fn int_expr(n: i64) -> Expr {
        Expr::Literal(CellValue::Int(n))
    }

    fn bool_expr(b: bool) -> Expr {
        Expr::Literal(CellValue::Bool(b))
    }

    // NORM.DIST tests
    #[tokio::test]
    async fn test_eval_norm_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(42.0),
            num_expr(40.0),
            num_expr(1.5),
            bool_expr(true),
        ];
        let result = eval_norm_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.9088).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_norm_dist_pdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(42.0),
            num_expr(40.0),
            num_expr(1.5),
            bool_expr(false),
        ];
        let result = eval_norm_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.1093).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_norm_dist_invalid_std_dev() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(42.0),
            num_expr(40.0),
            num_expr(0.0),
            bool_expr(true),
        ];
        let result = eval_norm_dist(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    #[tokio::test]
    async fn test_eval_norm_dist_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0), num_expr(40.0)];
        let result = eval_norm_dist(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    // NORM.S.INV tests
    #[tokio::test]
    async fn test_eval_norm_s_inv() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.95)];
        let result = eval_norm_s_inv(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.645).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_norm_s_inv_invalid_prob() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.5)];
        let result = eval_norm_s_inv(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    #[tokio::test]
    async fn test_eval_norm_s_inv_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_norm_s_inv(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    // NORM.S.DIST tests
    #[tokio::test]
    async fn test_eval_norm_s_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.5), bool_expr(true)];
        let result = eval_norm_s_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.9332).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_norm_s_dist_pdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0), bool_expr(false)];
        let result = eval_norm_s_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.3989).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_norm_s_dist_default() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_norm_s_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.5).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    // BETA.DIST tests
    #[tokio::test]
    async fn test_eval_beta_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.5), num_expr(2.0), num_expr(3.0), bool_expr(true)];
        let result = eval_beta_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.6875).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_beta_dist_scaled() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // With custom bounds A=0, B=2
        let args = vec![
            num_expr(1.0),
            num_expr(2.0),
            num_expr(3.0),
            bool_expr(true),
            num_expr(0.0),
            num_expr(2.0),
        ];
        let result = eval_beta_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.6875).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_beta_dist_invalid_alpha() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.5), num_expr(0.0), num_expr(3.0), bool_expr(true)];
        let result = eval_beta_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_beta_dist_x_out_of_bounds() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.5), num_expr(2.0), num_expr(3.0), bool_expr(true)];
        let result = eval_beta_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    // BETA.INV tests
    #[tokio::test]
    async fn test_eval_beta_inv() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.6875), num_expr(2.0), num_expr(3.0)];
        let result = eval_beta_inv(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.5).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_beta_inv_scaled() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(0.6875),
            num_expr(2.0),
            num_expr(3.0),
            num_expr(0.0),
            num_expr(2.0),
        ];
        let result = eval_beta_inv(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_beta_inv_invalid_prob() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-0.5), num_expr(2.0), num_expr(3.0)];
        let result = eval_beta_inv(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    // DEVSQ tests
    #[tokio::test]
    async fn test_eval_devsq() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(4));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(5));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(8));
        engine.set_cell("Sheet1", 3, 0, CellValue::Int(9));
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 3,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_devsq(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // Mean = (4+5+8+9)/4 = 6.5
            // DEVSQ = (-2.5)^2 + (-1.5)^2 + (1.5)^2 + (2.5)^2 = 6.25 + 2.25 + 2.25 + 6.25 = 17.0
            CellValue::Float(v) => assert!((v - 17.0).abs() < 0.001),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_devsq_no_values() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_devsq(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    // PROB tests
    #[tokio::test]
    async fn test_eval_prob() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Set up x_range
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(1));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(2));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(3));
        engine.set_cell("Sheet1", 3, 0, CellValue::Int(4));
        // Set up prob_range
        engine.set_cell("Sheet1", 0, 1, CellValue::Float(0.1));
        engine.set_cell("Sheet1", 1, 1, CellValue::Float(0.2));
        engine.set_cell("Sheet1", 2, 1, CellValue::Float(0.3));
        engine.set_cell("Sheet1", 3, 1, CellValue::Float(0.4));
        let x_range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 3,
            end_col: 0,
        });
        let prob_range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 1,
            end_row: 3,
            end_col: 1,
        });
        let args = vec![x_range, prob_range, num_expr(2.0), num_expr(3.0)];
        let result = eval_prob(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.5).abs() < 0.001),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    // CHISQ.DIST tests
    #[tokio::test]
    async fn test_eval_chisq_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(6.0), num_expr(3.0), bool_expr(true)];
        let result = eval_chisq_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.8884).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_chisq_dist_pdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(6.0), num_expr(3.0), bool_expr(false)];
        let result = eval_chisq_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.0),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_chisq_dist_invalid_df() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(6.0), num_expr(0.0), bool_expr(true)];
        let result = eval_chisq_dist(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    // CHISQ.DIST.RT tests
    #[tokio::test]
    async fn test_eval_chisq_dist_rt() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(6.0), num_expr(3.0)];
        let result = eval_chisq_dist_rt(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.1116).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_chisq_dist_rt_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(6.0)];
        let result = eval_chisq_dist_rt(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    // CHISQ.INV tests
    #[tokio::test]
    async fn test_eval_chisq_inv() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.5), num_expr(10.0)];
        let result = eval_chisq_inv(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 9.3418).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_chisq_inv_invalid() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.5), num_expr(10.0)];
        let result = eval_chisq_inv(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    // CHISQ.INV.RT tests
    #[tokio::test]
    async fn test_eval_chisq_inv_rt() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.5), num_expr(10.0)];
        let result = eval_chisq_inv_rt(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // CHISQ.INV.RT(0.5, 10) returns x where P(X > x) = 0.5, which equals CHISQ.INV(0.5, 10)
            CellValue::Float(v) => assert!((v - 9.3418).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    // T.DIST tests
    #[tokio::test]
    async fn test_eval_t_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.5), num_expr(10.0), bool_expr(true)];
        let result = eval_t_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.9177).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_t_dist_pdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.5), num_expr(10.0), bool_expr(false)];
        let result = eval_t_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.0),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_t_dist_invalid_df() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.5), num_expr(0.0), bool_expr(true)];
        let result = eval_t_dist(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    // T.DIST.2T tests
    #[tokio::test]
    async fn test_eval_t_dist_2t() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.5), num_expr(10.0)];
        let result = eval_t_dist_2t(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.1646).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    // T.DIST.RT tests
    #[tokio::test]
    async fn test_eval_t_dist_rt() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.5), num_expr(10.0)];
        let result = eval_t_dist_rt(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.0823).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    // T.INV tests
    #[tokio::test]
    async fn test_eval_t_inv() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.95), num_expr(10.0)];
        let result = eval_t_inv(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.8125).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_t_inv_invalid_prob() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-0.1), num_expr(10.0)];
        let result = eval_t_inv(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    // T.INV.2T tests
    #[tokio::test]
    async fn test_eval_t_inv_2t() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.1), num_expr(10.0)];
        let result = eval_t_inv_2t(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.8125).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    // F.DIST tests
    #[tokio::test]
    async fn test_eval_f_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(2.0),
            num_expr(5.0),
            num_expr(10.0),
            bool_expr(true),
        ];
        let result = eval_f_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.8358).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_f_dist_pdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(2.0),
            num_expr(5.0),
            num_expr(10.0),
            bool_expr(false),
        ];
        let result = eval_f_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.0),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_f_dist_invalid_df() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(2.0),
            num_expr(0.0),
            num_expr(10.0),
            bool_expr(true),
        ];
        let result = eval_f_dist(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    // F.DIST.RT tests
    #[tokio::test]
    async fn test_eval_f_dist_rt() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(2.0), num_expr(5.0), num_expr(10.0)];
        let result = eval_f_dist_rt(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.1642).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    // F.INV tests
    #[tokio::test]
    async fn test_eval_f_inv() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.95), num_expr(5.0), num_expr(10.0)];
        let result = eval_f_inv(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.3258).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_f_inv_invalid() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.5), num_expr(5.0), num_expr(10.0)];
        let result = eval_f_inv(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    // F.INV.RT tests
    #[tokio::test]
    async fn test_eval_f_inv_rt() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.05), num_expr(5.0), num_expr(10.0)];
        let result = eval_f_inv_rt(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.3258).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    // BINOM.DIST tests
    #[tokio::test]
    async fn test_eval_binom_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![int_expr(6), int_expr(10), num_expr(0.5), bool_expr(true)];
        let result = eval_binom_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.8281).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_binom_dist_pmf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![int_expr(6), int_expr(10), num_expr(0.5), bool_expr(false)];
        let result = eval_binom_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.2051).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_binom_dist_invalid() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![int_expr(6), int_expr(10), num_expr(1.5), bool_expr(true)];
        let result = eval_binom_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    // HYPGEOM.DIST tests
    #[tokio::test]
    async fn test_eval_hypgeom_dist_pmf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // sample_s=1, number_sample=4, population_s=8, number_pop=20
        let args = vec![
            int_expr(1),
            int_expr(4),
            int_expr(8),
            int_expr(20),
            bool_expr(false),
        ];
        let result = eval_hypgeom_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.3633).abs() < 0.001),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_hypgeom_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            int_expr(1),
            int_expr(4),
            int_expr(8),
            int_expr(20),
            bool_expr(true),
        ];
        let result = eval_hypgeom_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.4 && v < 0.5),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    // NEGBINOM.DIST tests
    #[tokio::test]
    async fn test_eval_negbinom_dist_pmf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // number_f=2, number_s=5, probability_s=0.5
        let args = vec![int_expr(2), int_expr(5), num_expr(0.5), bool_expr(false)];
        let result = eval_negbinom_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.1172).abs() < 0.001),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    // POISSON.DIST tests
    #[tokio::test]
    async fn test_eval_poisson_dist_pmf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![int_expr(3), num_expr(5.0), bool_expr(false)];
        let result = eval_poisson_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.1404).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_poisson_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![int_expr(3), num_expr(5.0), bool_expr(true)];
        let result = eval_poisson_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.2650).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    // CONFIDENCE.NORM tests
    #[tokio::test]
    async fn test_eval_confidence_norm() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.05), num_expr(2.5), num_expr(100.0)];
        let result = eval_confidence_norm(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.49).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_confidence_norm_invalid_alpha() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0), num_expr(2.5), num_expr(100.0)];
        let result = eval_confidence_norm(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    // CONFIDENCE.T tests
    #[tokio::test]
    async fn test_eval_confidence_t() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.05), num_expr(2.5), num_expr(100.0)];
        let result = eval_confidence_t(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.49 && v < 0.5),
            _ => panic!("Expected Float"),
        }
    }

    // EXPON.DIST tests
    #[tokio::test]
    async fn test_eval_expon_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.5), num_expr(2.0), bool_expr(true)];
        let result = eval_expon_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.6321).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_expon_dist_pdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.5), num_expr(2.0), bool_expr(false)];
        let result = eval_expon_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.7358).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_expon_dist_invalid_x() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-0.5), num_expr(2.0), bool_expr(true)];
        let result = eval_expon_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    // GAMMA.DIST tests
    #[tokio::test]
    async fn test_eval_gamma_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(5.0), num_expr(3.0), num_expr(2.0), bool_expr(true)];
        let result = eval_gamma_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.4562).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_gamma_dist_invalid_x() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(-1.0),
            num_expr(3.0),
            num_expr(2.0),
            bool_expr(true),
        ];
        let result = eval_gamma_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    // GAMMAINV tests
    #[tokio::test]
    async fn test_eval_gammainv() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.5), num_expr(3.0), num_expr(2.0)];
        let result = eval_gammainv(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 5.3481).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    // GAMMALN tests
    #[tokio::test]
    async fn test_eval_gammaln() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(5.0)];
        let result = eval_gammaln(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.1781).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_gammaln_invalid() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_gammaln(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    // LOGNORM.DIST tests
    #[tokio::test]
    async fn test_eval_lognorm_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(4.0), num_expr(1.0), num_expr(0.5), bool_expr(true)];
        let result = eval_lognorm_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.7801).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_lognorm_dist_invalid_x() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0), num_expr(1.0), num_expr(0.5), bool_expr(true)];
        let result = eval_lognorm_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    // LOGNORM.INV tests
    #[tokio::test]
    async fn test_eval_lognorm_inv() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.5), num_expr(1.0), num_expr(0.5)];
        let result = eval_lognorm_inv(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 2.7183).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    // WEIBULL.DIST tests
    #[tokio::test]
    async fn test_eval_weibull_dist_cdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(5.0), num_expr(2.0), num_expr(3.0), bool_expr(true)];
        let result = eval_weibull_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.9378).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_weibull_dist_pdf() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(5.0),
            num_expr(2.0),
            num_expr(3.0),
            bool_expr(false),
        ];
        let result = eval_weibull_dist(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.0),
            _ => panic!("Expected Float"),
        }
    }

    // Z.TEST tests
    #[tokio::test]
    async fn test_eval_z_test() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Set up data range
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(3));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(6));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(7));
        engine.set_cell("Sheet1", 3, 0, CellValue::Int(8));
        engine.set_cell("Sheet1", 4, 0, CellValue::Int(6));
        engine.set_cell("Sheet1", 5, 0, CellValue::Int(5));
        engine.set_cell("Sheet1", 6, 0, CellValue::Int(4));
        engine.set_cell("Sheet1", 7, 0, CellValue::Int(2));
        engine.set_cell("Sheet1", 8, 0, CellValue::Int(1));
        engine.set_cell("Sheet1", 9, 0, CellValue::Int(9));
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 9,
            end_col: 0,
        });
        let args = vec![range, num_expr(4.0)];
        let result = eval_z_test(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.0 && v < 1.0),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    // F.TEST tests
    #[tokio::test]
    async fn test_eval_f_test() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Set up data ranges
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(6));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(7));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(9));
        engine.set_cell("Sheet1", 3, 0, CellValue::Int(15));
        engine.set_cell("Sheet1", 4, 0, CellValue::Int(21));

        engine.set_cell("Sheet1", 0, 1, CellValue::Int(20));
        engine.set_cell("Sheet1", 1, 1, CellValue::Int(28));
        engine.set_cell("Sheet1", 2, 1, CellValue::Int(31));
        engine.set_cell("Sheet1", 3, 1, CellValue::Int(38));
        engine.set_cell("Sheet1", 4, 1, CellValue::Int(40));

        let range1 = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 4,
            end_col: 0,
        });
        let range2 = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 1,
            end_row: 4,
            end_col: 1,
        });
        let args = vec![range1, range2];
        let result = eval_f_test(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.0 && v <= 1.0),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    // CHISQ.TEST tests
    #[tokio::test]
    async fn test_eval_chisq_test() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Set up actual range (2x2)
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(58));
        engine.set_cell("Sheet1", 0, 1, CellValue::Int(35));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(11));
        engine.set_cell("Sheet1", 1, 1, CellValue::Int(25));
        // Set up expected range (2x2)
        engine.set_cell("Sheet1", 2, 0, CellValue::Float(45.35));
        engine.set_cell("Sheet1", 2, 1, CellValue::Float(47.65));
        engine.set_cell("Sheet1", 3, 0, CellValue::Float(23.65));
        engine.set_cell("Sheet1", 3, 1, CellValue::Float(24.85));

        let actual = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 1,
            end_col: 1,
        });
        let expected = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 2,
            start_col: 0,
            end_row: 3,
            end_col: 1,
        });
        let args = vec![actual, expected];
        let result = eval_chisq_test(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.0 && v < 1.0),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    // T.TEST tests
    #[tokio::test]
    async fn test_eval_t_test_paired() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Set up paired data
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(3));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(4));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(5));
        engine.set_cell("Sheet1", 0, 1, CellValue::Int(6));
        engine.set_cell("Sheet1", 1, 1, CellValue::Int(19));
        engine.set_cell("Sheet1", 2, 1, CellValue::Int(3));

        let range1 = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let range2 = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 1,
            end_row: 2,
            end_col: 1,
        });
        // type=1 for paired, tails=2 for two-tailed
        let args = vec![range1, range2, int_expr(2), int_expr(1)];
        let result = eval_t_test(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v >= 0.0 && v <= 1.0),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_t_test_invalid_tails() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(3));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(4));

        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 1,
            end_col: 0,
        });
        // tails=3 is invalid
        let args = vec![range.clone(), range, int_expr(3), int_expr(1)];
        let result = eval_t_test(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    // BINOM.INV tests
    #[tokio::test]
    async fn test_eval_binom_inv() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(6.0), num_expr(0.5), num_expr(0.75)];
        let result = eval_binom_inv(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert!(v >= 0 && v <= 6),
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    // BINOM.DIST.RANGE tests
    #[tokio::test]
    async fn test_eval_binom_dist_range() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 6 trials, p=0.5, probability of exactly 4 successes
        let args = vec![num_expr(6.0), num_expr(0.5), num_expr(4.0)];
        let result = eval_binom_dist_range(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.2344).abs() < 0.001),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_binom_dist_range_between() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 6 trials, p=0.5, probability of between 2 and 4 successes (inclusive)
        let args = vec![num_expr(6.0), num_expr(0.5), num_expr(2.0), num_expr(4.0)];
        let result = eval_binom_dist_range(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v > 0.6 && v < 0.8),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    // NORM.INV tests
    #[tokio::test]
    async fn test_eval_norm_inv() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.95), num_expr(40.0), num_expr(1.5)];
        let result = eval_norm_inv(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 42.4673).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_norm_inv_invalid_prob() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-0.1), num_expr(40.0), num_expr(1.5)];
        let result = eval_norm_inv(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    // GAUSS tests
    #[tokio::test]
    async fn test_eval_gauss() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(2.0)];
        let result = eval_gauss(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.4772).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    // PHI tests
    #[tokio::test]
    async fn test_eval_phi() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_phi(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.3989).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_phi_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0), num_expr(1.0)];
        let result = eval_phi(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }
}
