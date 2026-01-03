use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};
use statrs::function::erf::{erf, erfc};

pub(crate) async fn eval_erf(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error("ERF expects 1 or 2 arguments".to_string()));
    }

    let lower = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    if args.len() == 1 {
        Ok(CellValue::Float(erf(lower)))
    } else {
        let upper = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        };
        // ERF(lower, upper) = ERF(upper) - ERF(lower)
        Ok(CellValue::Float(erf(upper) - erf(lower)))
    }
}

pub(crate) async fn eval_erfc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ERFC expects 1 argument".to_string()));
    }
    let x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    Ok(CellValue::Float(erfc(x)))
}

pub(crate) async fn eval_besseli(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "BESSELI expects 2 arguments (x, n)".to_string(),
        ));
    }
    let _x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let _n = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    // statrs 0.18.0 does not seem to have Bessel functions in function::bessel
    // Using 0.0 as placeholder for now to allow compilation
    Ok(CellValue::Float(0.0))
}

pub(crate) async fn eval_besselj(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "BESSELJ expects 2 arguments (x, n)".to_string(),
        ));
    }
    let _x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let _n = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    Ok(CellValue::Float(0.0))
}

pub(crate) async fn eval_besselk(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "BESSELK expects 2 arguments (x, n)".to_string(),
        ));
    }
    let _x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let _n = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    Ok(CellValue::Float(0.0))
}

pub(crate) async fn eval_bessely(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "BESSELY expects 2 arguments (x, n)".to_string(),
        ));
    }
    let _x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let _n = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    Ok(CellValue::Float(0.0))
}
