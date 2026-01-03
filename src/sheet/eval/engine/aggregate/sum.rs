use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, flatten_range_expr, for_each_value_in_expr, to_number};

pub(crate) async fn eval_sum(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut total = 0.0f64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                total += n;
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_product(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut product = 1.0f64;
    let mut found_numeric = false;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                product *= n;
                found_numeric = true;
            }
            Ok(())
        })
        .await?;
    }
    if !found_numeric {
        Ok(CellValue::Float(0.0))
    } else {
        Ok(CellValue::Float(product))
    }
}

pub(crate) async fn eval_sumproduct(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "SUMPRODUCT expects at least 1 argument".to_string(),
        ));
    }

    let mut flattened_args = Vec::with_capacity(args.len());
    for arg in args {
        flattened_args.push(flatten_range_expr(ctx, current_sheet, arg).await?);
    }

    // All arrays must have the same dimensions
    let rows = flattened_args[0].rows;
    let cols = flattened_args[0].cols;
    for arg in &flattened_args[1..] {
        if arg.rows != rows || arg.cols != cols {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }
    }

    let mut total = 0.0;
    for i in 0..(rows * cols) {
        let mut product = 1.0;
        for arg in &flattened_args {
            let val = &arg.values[i];
            let n = to_number(val).unwrap_or(0.0);
            product *= n;
        }
        total += product;
    }

    Ok(CellValue::Float(total))
}
