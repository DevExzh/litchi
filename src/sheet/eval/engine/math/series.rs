use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, for_each_value_in_expr, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};
use rand::Rng;

use super::helpers::flatten_numeric_values;

pub(crate) async fn eval_sumsq(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut total = 0.0f64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |value| {
            if let Some(n) = to_number(value) {
                total += n * n;
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_sumx2my2(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "SUMX2MY2 expects 2 arguments (array_x, array_y)".to_string(),
        ));
    }
    let x_vals = flatten_numeric_values(ctx, current_sheet, &args[0]).await?;
    let y_vals = flatten_numeric_values(ctx, current_sheet, &args[1]).await?;
    if x_vals.len() != y_vals.len() {
        return Ok(CellValue::Error(
            "SUMX2MY2 requires arrays of the same size".to_string(),
        ));
    }
    let total = x_vals
        .iter()
        .zip(y_vals.iter())
        .fold(0.0, |acc, (&x, &y)| acc + x * x - y * y);
    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_sumx2py2(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "SUMX2PY2 expects 2 arguments (array_x, array_y)".to_string(),
        ));
    }
    let x_vals = flatten_numeric_values(ctx, current_sheet, &args[0]).await?;
    let y_vals = flatten_numeric_values(ctx, current_sheet, &args[1]).await?;
    if x_vals.len() != y_vals.len() {
        return Ok(CellValue::Error(
            "SUMX2PY2 requires arrays of the same size".to_string(),
        ));
    }
    let total = x_vals
        .iter()
        .zip(y_vals.iter())
        .fold(0.0, |acc, (&x, &y)| acc + x * x + y * y);
    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_sumxmy2(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "SUMXMY2 expects 2 arguments (array_x, array_y)".to_string(),
        ));
    }
    let x_vals = flatten_numeric_values(ctx, current_sheet, &args[0]).await?;
    let y_vals = flatten_numeric_values(ctx, current_sheet, &args[1]).await?;
    if x_vals.len() != y_vals.len() {
        return Ok(CellValue::Error(
            "SUMXMY2 requires arrays of the same size".to_string(),
        ));
    }
    let total = x_vals.iter().zip(y_vals.iter()).fold(0.0, |acc, (&x, &y)| {
        let diff = x - y;
        acc + diff * diff
    });
    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_seriessum(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 4 {
        return Ok(CellValue::Error(
            "SERIESSUM expects 4 arguments (x, n, m, coefficients)".to_string(),
        ));
    }

    let x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let n = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let m = match to_number(&evaluate_expression(ctx, current_sheet, &args[2]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let mut coefficients = Vec::new();
    for_each_value_in_expr(ctx, current_sheet, &args[3], |val| {
        if let Some(c) = to_number(val) {
            coefficients.push(c);
        }
        Ok(())
    })
    .await?;

    if coefficients.is_empty() {
        return Ok(CellValue::Float(0.0));
    }

    let mut total = 0.0;
    for (i, &coeff) in coefficients.iter().enumerate() {
        let power = n + (i as f64) * m;
        total += coeff * x.powf(power);
    }

    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_sequence(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 4 {
        return Ok(CellValue::Error(
            "SEQUENCE expects 1 to 4 arguments (rows, [columns], [start], [step])".to_string(),
        ));
    }

    let rows = match args.first() {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) if n >= 0.0 => n.trunc() as usize,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1,
    };

    let cols = match args.get(1) {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) if n >= 0.0 => n.trunc() as usize,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1,
    };

    let start = match args.get(2) {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) => n,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1.0,
    };

    let step = match args.get(3) {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) => n,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1.0,
    };

    if rows == 0 || cols == 0 {
        return Ok(CellValue::Error("#CALC!".to_string()));
    }

    let _step = step; // Avoid unused warning for now
    // Since CellValue doesn't support arrays yet, return the first value (start)
    Ok(CellValue::Float(start))
}

pub(crate) async fn eval_vstack(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "VSTACK expects at least 1 argument".to_string(),
        ));
    }
    // Placeholder: return first value of first argument
    evaluate_expression(ctx, current_sheet, &args[0]).await
}

pub(crate) async fn eval_hstack(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "HSTACK expects at least 1 argument".to_string(),
        ));
    }
    // Placeholder: return first value of first argument
    evaluate_expression(ctx, current_sheet, &args[0]).await
}

pub(crate) async fn eval_wrapcols(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "WRAPCOLS expects 2 or 3 arguments".to_string(),
        ));
    }
    // Placeholder: return first value of first argument
    evaluate_expression(ctx, current_sheet, &args[0]).await
}

pub(crate) async fn eval_wraprows(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "WRAPROWS expects 2 or 3 arguments".to_string(),
        ));
    }
    // Placeholder: return first value of first argument
    evaluate_expression(ctx, current_sheet, &args[0]).await
}

pub(crate) async fn eval_randarray(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() > 5 {
        return Ok(CellValue::Error(
            "RANDARRAY expects 0 to 5 arguments (rows, [columns], [min], [max], [whole_number])"
                .to_string(),
        ));
    }

    let rows = match args.first() {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) if n >= 0.0 => n.trunc() as usize,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1,
    };

    let cols = match args.get(1) {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) if n >= 0.0 => n.trunc() as usize,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1,
    };

    let min = match args.get(2) {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) => n,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 0.0,
    };

    let max = match args.get(3) {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) => n,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1.0,
    };

    let whole_number = match args.get(4) {
        Some(expr) => {
            let val = evaluate_expression(ctx, current_sheet, expr).await?;
            match val {
                CellValue::Bool(b) => b,
                _ => match to_number(&val) {
                    Some(n) => n != 0.0,
                    None => false,
                },
            }
        },
        None => false,
    };

    if rows == 0 || cols == 0 {
        return Ok(CellValue::Error("#CALC!".to_string()));
    }

    if min > max {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }

    let mut rng = rand::rng();
    let val = if whole_number {
        let bottom = min.ceil() as i64;
        let top = max.floor() as i64;
        if bottom > top {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }
        CellValue::Int(rng.random_range(bottom..=top))
    } else {
        CellValue::Float(min + (max - min) * rng.random::<f64>())
    };

    // Since CellValue doesn't support arrays yet, return a single random value
    Ok(val)
}
