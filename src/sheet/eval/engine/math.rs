use crate::sheet::{CellValue, Result};

use super::super::{EngineCtx, parser::Expr};
use super::{evaluate_expression, to_number};

pub(crate) fn eval_int<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("INT expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0])?;
    let n = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("INT on non-numeric value".to_string())),
    };
    Ok(CellValue::Int(n.floor() as i64))
}

pub(crate) fn eval_abs<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ABS expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0])?;
    match v {
        CellValue::Int(i) => Ok(CellValue::Int(i.abs())),
        CellValue::Float(f) => Ok(CellValue::Float(f.abs())),
        other => Ok(CellValue::Error(format!(
            "ABS on non-numeric value: {:?}",
            other
        ))),
    }
}

pub(crate) fn eval_power<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("POWER expects 2 arguments".to_string()));
    }
    let base = evaluate_expression(ctx, current_sheet, &args[0])?;
    let exp = evaluate_expression(ctx, current_sheet, &args[1])?;
    let b = match to_number(&base) {
        Some(n) => n,
        None => return Ok(CellValue::Error("POWER base is not numeric".to_string())),
    };
    let e = match to_number(&exp) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "POWER exponent is not numeric".to_string(),
            ));
        },
    };
    Ok(CellValue::Float(b.powf(e)))
}

pub(crate) fn eval_round<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("ROUND expects 2 arguments".to_string()));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0])?;
    let digits = evaluate_expression(ctx, current_sheet, &args[1])?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => return Ok(CellValue::Error("ROUND number is not numeric".to_string())),
    };
    let d = match to_number(&digits) {
        Some(n) => n as i32,
        None => return Ok(CellValue::Error("ROUND digits is not numeric".to_string())),
    };
    let factor = 10f64.powi(d.abs());
    let result = if d >= 0 {
        (x * factor).round() / factor
    } else {
        (x / factor).round() * factor
    };
    Ok(CellValue::Float(result))
}

pub(crate) fn eval_rounddown<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "ROUNDDOWN expects 2 arguments".to_string(),
        ));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0])?;
    let digits = evaluate_expression(ctx, current_sheet, &args[1])?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "ROUNDDOWN number is not numeric".to_string(),
            ));
        },
    };
    let d = match to_number(&digits) {
        Some(n) => n as i32,
        None => {
            return Ok(CellValue::Error(
                "ROUNDDOWN digits is not numeric".to_string(),
            ));
        },
    };
    let factor = 10f64.powi(d.abs());
    let result = if d >= 0 {
        (x * factor).trunc() / factor
    } else {
        (x / factor).trunc() * factor
    };
    Ok(CellValue::Float(result))
}

pub(crate) fn eval_roundup<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("ROUNDUP expects 2 arguments".to_string()));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0])?;
    let digits = evaluate_expression(ctx, current_sheet, &args[1])?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "ROUNDUP number is not numeric".to_string(),
            ));
        },
    };
    let d = match to_number(&digits) {
        Some(n) => n as i32,
        None => {
            return Ok(CellValue::Error(
                "ROUNDUP digits is not numeric".to_string(),
            ));
        },
    };
    let factor = 10f64.powi(d.abs());
    let scaled = if d >= 0 { x * factor } else { x / factor };
    let rounded = if scaled >= 0.0 {
        scaled.ceil()
    } else {
        scaled.floor()
    };
    let result = if d >= 0 {
        rounded / factor
    } else {
        rounded * factor
    };
    Ok(CellValue::Float(result))
}

pub(crate) fn eval_floor<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "FLOOR expects 1 or 2 arguments (number, [significance])".to_string(),
        ));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0])?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => return Ok(CellValue::Error("FLOOR number is not numeric".to_string())),
    };
    let sig = if args.len() == 2 {
        let s_val = evaluate_expression(ctx, current_sheet, &args[1])?;
        match to_number(&s_val) {
            Some(n) if n != 0.0 => n,
            _ => {
                return Ok(CellValue::Error(
                    "FLOOR significance must be non-zero numeric".to_string(),
                ));
            },
        }
    } else {
        1.0
    };
    let result = (x / sig).floor() * sig;
    Ok(CellValue::Float(result))
}

pub(crate) fn eval_ceiling<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "CEILING expects 1 or 2 arguments (number, [significance])".to_string(),
        ));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0])?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "CEILING number is not numeric".to_string(),
            ));
        },
    };
    let sig = if args.len() == 2 {
        let s_val = evaluate_expression(ctx, current_sheet, &args[1])?;
        match to_number(&s_val) {
            Some(n) if n != 0.0 => n,
            _ => {
                return Ok(CellValue::Error(
                    "CEILING significance must be non-zero numeric".to_string(),
                ));
            },
        }
    } else {
        1.0
    };
    let result = (x / sig).ceil() * sig;
    Ok(CellValue::Float(result))
}

fn eval_unary_math<C, F>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
    fname: &str,
    op: F,
) -> Result<CellValue>
where
    C: EngineCtx + ?Sized,
    F: Fn(f64) -> f64,
{
    if args.len() != 1 {
        return Ok(CellValue::Error(format!("{} expects 1 argument", fname)));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0])?;
    let n = match to_number(&v) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(format!(
                "{} on non-numeric value: {:?}",
                fname, v
            )));
        },
    };
    Ok(CellValue::Float(op(n)))
}

pub(crate) fn eval_sqrt<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "SQRT", |x| x.sqrt())
}

pub(crate) fn eval_exp<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "EXP", |x| x.exp())
}

pub(crate) fn eval_ln<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "LN", |x| x.ln())
}

pub(crate) fn eval_log10<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "LOG10", |x| x.log10())
}

pub(crate) fn eval_sin<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "SIN", |x| x.sin())
}

pub(crate) fn eval_cos<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "COS", |x| x.cos())
}

pub(crate) fn eval_tan<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "TAN", |x| x.tan())
}

pub(crate) fn eval_asin<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "ASIN", |x| x.asin())
}

pub(crate) fn eval_acos<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "ACOS", |x| x.acos())
}

pub(crate) fn eval_atan<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "ATAN", |x| x.atan())
}

pub(crate) fn eval_atan2<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("ATAN2 expects 2 arguments".to_string()));
    }
    let y = evaluate_expression(ctx, current_sheet, &args[0])?;
    let x = evaluate_expression(ctx, current_sheet, &args[1])?;
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
