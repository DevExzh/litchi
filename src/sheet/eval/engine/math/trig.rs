use std::f64::consts::FRAC_PI_2;

use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::helpers::EPS;

pub(crate) async fn eval_sin(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "SIN", f64::sin).await
}

pub(crate) async fn eval_cos(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "COS", f64::cos).await
}

pub(crate) async fn eval_tan(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "TAN", f64::tan).await
}

pub(crate) async fn eval_cot(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("COT expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let angle = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("COT on non-numeric value".to_string())),
    };
    let tan = angle.tan();
    if tan.abs() < EPS {
        return Ok(CellValue::Error(
            "COT is undefined for angles where TAN(angle) = 0".to_string(),
        ));
    }
    Ok(CellValue::Float(1.0 / tan))
}

pub(crate) async fn eval_acot(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ACOT", |x| FRAC_PI_2 - x.atan()).await
}

pub(crate) async fn eval_csc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("CSC expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let angle = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("CSC on non-numeric value".to_string())),
    };
    let sin = angle.sin();
    if sin.abs() < EPS {
        return Ok(CellValue::Error(
            "CSC is undefined for angles where SIN(angle) = 0".to_string(),
        ));
    }
    Ok(CellValue::Float(1.0 / sin))
}

pub(crate) async fn eval_sec(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("SEC expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let angle = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("SEC on non-numeric value".to_string())),
    };
    let cos = angle.cos();
    if cos.abs() < EPS {
        return Ok(CellValue::Error(
            "SEC is undefined for angles where COS(angle) = 0".to_string(),
        ));
    }
    Ok(CellValue::Float(1.0 / cos))
}

pub(crate) async fn eval_sinh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "SINH", f64::sinh).await
}

pub(crate) async fn eval_cosh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "COSH", f64::cosh).await
}

pub(crate) async fn eval_tanh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "TANH", f64::tanh).await
}

pub(crate) async fn eval_csch(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("CSCH expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let value = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("CSCH on non-numeric value".to_string())),
    };
    let sinh = value.sinh();
    if sinh.abs() < EPS {
        return Ok(CellValue::Error(
            "CSCH is undefined for values where SINH(value) = 0".to_string(),
        ));
    }
    Ok(CellValue::Float(1.0 / sinh))
}

pub(crate) async fn eval_sech(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("SECH expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let value = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("SECH on non-numeric value".to_string())),
    };
    let cosh = value.cosh();
    if cosh.abs() < EPS {
        return Ok(CellValue::Error(
            "SECH is undefined for values where COSH(value) = 0".to_string(),
        ));
    }
    Ok(CellValue::Float(1.0 / cosh))
}

pub(crate) async fn eval_coth(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("COTH expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let value = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("COTH on non-numeric value".to_string())),
    };
    let sinh = value.sinh();
    if sinh.abs() < EPS {
        return Ok(CellValue::Error(
            "COTH is undefined for values where SINH(value) = 0".to_string(),
        ));
    }
    Ok(CellValue::Float(value.cosh() / sinh))
}

pub(crate) async fn eval_asin(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ASIN", f64::asin).await
}

pub(crate) async fn eval_acos(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ACOS", f64::acos).await
}

pub(crate) async fn eval_atan(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ATAN", f64::atan).await
}

pub(crate) async fn eval_atan2(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("ATAN2 expects 2 arguments".to_string()));
    }
    let y = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let yn = match to_number(&y) {
        Some(n) => n,
        None => return Ok(CellValue::Error("ATAN2 y is not numeric".to_string())),
    };
    let xn = match to_number(&x) {
        Some(n) => n,
        None => return Ok(CellValue::Error("ATAN2 x is not numeric".to_string())),
    };
    Ok(CellValue::Float(yn.atan2(xn)))
}

pub(crate) async fn eval_asinh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ASINH", f64::asinh).await
}

pub(crate) async fn eval_acosh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ACOSH", f64::acosh).await
}

pub(crate) async fn eval_atanh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ATANH", f64::atanh).await
}

pub(crate) async fn eval_acoth(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ACOTH expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = match to_number(&v) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "ACOTH requires a numeric argument".to_string(),
            ));
        },
    };
    if x.abs() <= 1.0 {
        return Ok(CellValue::Error(
            "ACOTH input must satisfy |x| > 1".to_string(),
        ));
    }
    let value = 0.5 * ((x + 1.0) / (x - 1.0)).ln();
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_radians(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "RADIANS", f64::to_radians).await
}

pub(crate) async fn eval_degrees(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "DEGREES", f64::to_degrees).await
}

async fn eval_unary_trig<F>(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    name: &str,
    f: F,
) -> Result<CellValue>
where
    F: Fn(f64) -> f64,
{
    if args.len() != 1 {
        return Ok(CellValue::Error(format!("{name} expects 1 argument")));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let n = match to_number(&v) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(format!(
                "{name} on non-numeric value: {:?}",
                v
            )));
        },
    };
    Ok(CellValue::Float(f(n)))
}
