use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, evaluate_expression, flatten_range_expr, to_number};
use super::helpers::{find_exact_match_index, is_1d};

pub(crate) async fn eval_xlookup(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 6 {
        return Ok(CellValue::Error(
            "XLOOKUP expects 3 to 6 arguments (lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])"
                .to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let lookup_array = flatten_range_expr(ctx, current_sheet, &args[1]).await?;
    let return_array = flatten_range_expr(ctx, current_sheet, &args[2]).await?;

    if !is_1d(&lookup_array) || !is_1d(&return_array) {
        return Ok(CellValue::Error(
            "XLOOKUP lookup_array and return_array must be one-dimensional ranges".to_string(),
        ));
    }

    if lookup_array.values.len() != return_array.values.len() {
        return Ok(CellValue::Error(
            "XLOOKUP lookup_array and return_array must have the same length".to_string(),
        ));
    }

    let if_not_found = if args.len() >= 4 {
        Some(evaluate_expression(ctx, current_sheet, &args[3]).await?)
    } else {
        None
    };

    if args.len() >= 5 {
        let mm_val = evaluate_expression(ctx, current_sheet, &args[4]).await?;
        match to_number(&mm_val) {
            Some(0.0) => {},
            _ => {
                return Ok(CellValue::Error(
                    "XLOOKUP currently only supports match_mode = 0 (exact match)".to_string(),
                ));
            },
        }
    }

    if args.len() == 6 {
        return Ok(CellValue::Error(
            "XLOOKUP search_mode is not supported in this evaluator".to_string(),
        ));
    }

    if let Some(idx) = find_exact_match_index(&lookup_val, &lookup_array.values) {
        Ok(return_array.values[idx].clone())
    } else if let Some(not_found) = if_not_found {
        Ok(not_found)
    } else {
        Ok(CellValue::Error("XLOOKUP: value not found".to_string()))
    }
}
