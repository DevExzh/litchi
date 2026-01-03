use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, evaluate_expression, flatten_range_expr, to_bool, to_number};
use super::helpers::values_equal;

pub(crate) async fn eval_vlookup(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "VLOOKUP expects 3 or 4 arguments (lookup_value, table_array, col_index_num, [range_lookup])"
                .to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let table = flatten_range_expr(ctx, current_sheet, &args[1]).await?;

    let col_index_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
    let col_index = match to_number(&col_index_val) {
        Some(n) if n >= 1.0 => n as i64,
        _ => {
            return Ok(CellValue::Error(
                "VLOOKUP col_index_num must be a positive number".to_string(),
            ));
        },
    };

    let exact_match_only = if args.len() == 4 {
        let rl_val = evaluate_expression(ctx, current_sheet, &args[3]).await?;
        !to_bool(&rl_val)
    } else {
        true
    };

    if !exact_match_only {
        return Ok(CellValue::Error(
            "VLOOKUP currently only supports exact match (range_lookup = FALSE)".to_string(),
        ));
    }

    let rows = table.rows as i64;
    let cols = table.cols as i64;

    if col_index < 1 || col_index > cols {
        return Ok(CellValue::Error(
            "VLOOKUP col_index_num out of bounds for table_array".to_string(),
        ));
    }

    for r in 0..rows {
        let base = (r * cols) as usize;
        let key = &table.values[base];
        if values_equal(&lookup_val, key) {
            let idx = base + (col_index - 1) as usize;
            return Ok(table.values[idx].clone());
        }
    }

    Ok(CellValue::Error("VLOOKUP: value not found".to_string()))
}

pub(crate) async fn eval_hlookup(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "HLOOKUP expects 3 or 4 arguments (lookup_value, table_array, row_index_num, [range_lookup])"
                .to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let table = flatten_range_expr(ctx, current_sheet, &args[1]).await?;

    let row_index_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
    let row_index = match to_number(&row_index_val) {
        Some(n) if n >= 1.0 => n as i64,
        _ => {
            return Ok(CellValue::Error(
                "HLOOKUP row_index_num must be a positive number".to_string(),
            ));
        },
    };

    let exact_match_only = if args.len() == 4 {
        let rl_val = evaluate_expression(ctx, current_sheet, &args[3]).await?;
        !to_bool(&rl_val)
    } else {
        true
    };

    if !exact_match_only {
        return Ok(CellValue::Error(
            "HLOOKUP currently only supports exact match (range_lookup = FALSE)".to_string(),
        ));
    }

    let rows = table.rows as i64;
    let cols = table.cols as i64;

    if row_index < 1 || row_index > rows {
        return Ok(CellValue::Error(
            "HLOOKUP row_index_num out of bounds for table_array".to_string(),
        ));
    }

    for c in 0..cols {
        let key = &table.values[c as usize];
        if values_equal(&lookup_val, key) {
            let idx = ((row_index - 1) * cols + c) as usize;
            return Ok(table.values[idx].clone());
        }
    }

    Ok(CellValue::Error("HLOOKUP: value not found".to_string()))
}
