use crate::sheet::{CellValue, Result};

use super::super::parser::Expr;
use super::{EngineCtx, evaluate_expression, flatten_range_expr, to_bool, to_number};
use statrs::distribution::{
    ChiSquared, Continuous, ContinuousCDF, FisherSnedecor, Normal, StudentsT,
};

pub(crate) fn eval_norm_dist<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "NORM.DIST expects 3 or 4 arguments (x, mean, standard_dev, [cumulative])".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("NORM.DIST x is not numeric".to_string()));
        },
    };
    let mean = match number_arg(ctx, current_sheet, &args[1])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "NORM.DIST mean is not numeric".to_string(),
            ));
        },
    };
    let std_dev = match number_arg(ctx, current_sheet, &args[2])? {
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
        let v = evaluate_expression(ctx, current_sheet, &args[3])?;
        to_bool(&v)
    } else {
        true
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) fn eval_norm_s_inv<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error(
            "NORM.S.INV expects 1 argument (probability)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0])? {
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

pub(crate) fn eval_prob<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "PROB expects 3 or 4 arguments (x_range, prob_range, lower_limit, [upper_limit])"
                .to_string(),
        ));
    }

    let x_range = flatten_range_expr(ctx, current_sheet, &args[0])?;
    let p_range = flatten_range_expr(ctx, current_sheet, &args[1])?;

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

    let lower = match number_arg(ctx, current_sheet, &args[2])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "PROB lower_limit is not numeric".to_string(),
            ));
        },
    };

    let (lo, hi) = if args.len() == 4 {
        let upper = match number_arg(ctx, current_sheet, &args[3])? {
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

pub(crate) fn eval_chisq_dist<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "CHISQ.DIST expects 2 or 3 arguments (x, deg_freedom, [cumulative])".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("CHISQ.DIST x is not numeric".to_string()));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1])? {
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
        let v = evaluate_expression(ctx, current_sheet, &args[2])?;
        to_bool(&v)
    } else {
        true
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) fn eval_chisq_dist_rt<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "CHISQ.DIST.RT expects 2 arguments (x, deg_freedom)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "CHISQ.DIST.RT x is not numeric".to_string(),
            ));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1])? {
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

pub(crate) fn eval_chisq_inv<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "CHISQ.INV expects 2 arguments (probability, deg_freedom)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "CHISQ.INV probability is not numeric".to_string(),
            ));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1])? {
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

pub(crate) fn eval_chisq_inv_rt<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "CHISQ.INV.RT expects 2 arguments (probability, deg_freedom)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "CHISQ.INV.RT probability is not numeric".to_string(),
            ));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1])? {
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

pub(crate) fn eval_t_dist<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "T.DIST expects 3 arguments (x, deg_freedom, cumulative)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("T.DIST x is not numeric".to_string()));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1])? {
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
        let v = evaluate_expression(ctx, current_sheet, &args[2])?;
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

pub(crate) fn eval_t_dist_2t<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "T.DIST.2T expects 2 arguments (x, deg_freedom)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("T.DIST.2T x is not numeric".to_string()));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1])? {
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

pub(crate) fn eval_t_dist_rt<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "T.DIST.RT expects 2 arguments (x, deg_freedom)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("T.DIST.RT x is not numeric".to_string()));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1])? {
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

pub(crate) fn eval_t_inv<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "T.INV expects 2 arguments (probability, deg_freedom)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "T.INV probability is not numeric".to_string(),
            ));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1])? {
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

pub(crate) fn eval_t_inv_2t<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "T.INV.2T expects 2 arguments (probability, deg_freedom)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "T.INV.2T probability is not numeric".to_string(),
            ));
        },
    };
    let df = match number_arg(ctx, current_sheet, &args[1])? {
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

pub(crate) fn eval_f_dist<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "F.DIST expects 3 or 4 arguments (x, deg_freedom1, deg_freedom2, [cumulative])"
                .to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("F.DIST x is not numeric".to_string()));
        },
    };
    let df1 = match number_arg(ctx, current_sheet, &args[1])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.DIST deg_freedom1 is not numeric".to_string(),
            ));
        },
    };
    let df2 = match number_arg(ctx, current_sheet, &args[2])? {
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
        let v = evaluate_expression(ctx, current_sheet, &args[3])?;
        to_bool(&v)
    } else {
        true
    };

    let value = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    Ok(CellValue::Float(value))
}

pub(crate) fn eval_f_dist_rt<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "F.DIST.RT expects 3 arguments (x, deg_freedom1, deg_freedom2)".to_string(),
        ));
    }

    let x = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error("F.DIST.RT x is not numeric".to_string()));
        },
    };
    let df1 = match number_arg(ctx, current_sheet, &args[1])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.DIST.RT deg_freedom1 is not numeric".to_string(),
            ));
        },
    };
    let df2 = match number_arg(ctx, current_sheet, &args[2])? {
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

pub(crate) fn eval_f_inv<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "F.INV expects 3 arguments (probability, deg_freedom1, deg_freedom2)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.INV probability is not numeric".to_string(),
            ));
        },
    };
    let df1 = match number_arg(ctx, current_sheet, &args[1])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.INV deg_freedom1 is not numeric".to_string(),
            ));
        },
    };
    let df2 = match number_arg(ctx, current_sheet, &args[2])? {
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

pub(crate) fn eval_f_inv_rt<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "F.INV.RT expects 3 arguments (probability, deg_freedom1, deg_freedom2)".to_string(),
        ));
    }

    let p = match number_arg(ctx, current_sheet, &args[0])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.INV.RT probability is not numeric".to_string(),
            ));
        },
    };
    let df1 = match number_arg(ctx, current_sheet, &args[1])? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "F.INV.RT deg_freedom1 is not numeric".to_string(),
            ));
        },
    };
    let df2 = match number_arg(ctx, current_sheet, &args[2])? {
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

fn number_arg<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    expr: &Expr,
) -> Result<Option<f64>> {
    let v = evaluate_expression(ctx, current_sheet, expr)?;
    Ok(to_number(&v))
}
