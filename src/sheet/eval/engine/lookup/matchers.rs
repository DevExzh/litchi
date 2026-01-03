use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, evaluate_expression, flatten_range_expr, to_number};
use super::helpers::{find_exact_match_index, is_1d};

pub(crate) async fn eval_match(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "MATCH expects 2 or 3 arguments (lookup_value, lookup_array, [match_type])".to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let lookup_array = flatten_range_expr(ctx, current_sheet, &args[1]).await?;

    if !is_1d(&lookup_array) {
        return Ok(CellValue::Error(
            "MATCH lookup_array must be a one-dimensional range".to_string(),
        ));
    }

    let match_type = if args.len() == 3 {
        let mt_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match to_number(&mt_val) {
            Some(0.0) => 0,
            _ => {
                return Ok(CellValue::Error(
                    "MATCH currently only supports match_type = 0 (exact match)".to_string(),
                ));
            },
        }
    } else {
        0
    };

    if match_type != 0 {
        return Ok(CellValue::Error(
            "MATCH currently only supports exact match (match_type = 0)".to_string(),
        ));
    }

    if let Some(idx) = find_exact_match_index(&lookup_val, &lookup_array.values) {
        Ok(CellValue::Int((idx + 1) as i64))
    } else {
        Ok(CellValue::Error("MATCH: value not found".to_string()))
    }
}

pub(crate) async fn eval_xmatch(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 4 {
        return Ok(CellValue::Error(
            "XMATCH expects 2 to 4 arguments (lookup_value, lookup_array, [match_mode], [search_mode])"
                .to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let lookup_array = flatten_range_expr(ctx, current_sheet, &args[1]).await?;

    if !is_1d(&lookup_array) {
        return Ok(CellValue::Error(
            "XMATCH lookup_array must be a one-dimensional range".to_string(),
        ));
    }

    if args.len() >= 3 {
        let mm_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match to_number(&mm_val) {
            Some(0.0) => {},
            _ => {
                return Ok(CellValue::Error(
                    "XMATCH currently only supports match_mode = 0 (exact match)".to_string(),
                ));
            },
        }
    }

    if args.len() == 4 {
        return Ok(CellValue::Error(
            "XMATCH search_mode is not supported in this evaluator".to_string(),
        ));
    }

    if let Some(idx) = find_exact_match_index(&lookup_val, &lookup_array.values) {
        Ok(CellValue::Int((idx + 1) as i64))
    } else {
        Ok(CellValue::Error("XMATCH: value not found".to_string()))
    }
}
