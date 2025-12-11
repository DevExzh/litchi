use crate::sheet::{CellValue, Result};

use super::super::parser::Expr;
use super::{
    EngineCtx, FlatRange, evaluate_expression, flatten_range_expr, to_bool, to_number, to_text,
};

pub(crate) fn eval_index<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "INDEX expects 2 or 3 arguments (array, row_num, [column_num])".to_string(),
        ));
    }

    let array = flatten_range_expr(ctx, current_sheet, &args[0])?;

    let row_val = evaluate_expression(ctx, current_sheet, &args[1])?;
    let row_num = match to_number(&row_val) {
        Some(n) if n >= 1.0 => n as i64,
        _ => {
            return Ok(CellValue::Error(
                "INDEX row_num must be a positive number".to_string(),
            ));
        },
    };

    let col_num = if args.len() == 3 {
        let col_val = evaluate_expression(ctx, current_sheet, &args[2])?;
        match to_number(&col_val) {
            Some(n) if n >= 1.0 => n as i64,
            _ => {
                return Ok(CellValue::Error(
                    "INDEX column_num must be a positive number".to_string(),
                ));
            },
        }
    } else {
        1
    };

    let rows = array.rows as i64;
    let cols = array.cols as i64;

    if row_num < 1 || row_num > rows || col_num < 1 || col_num > cols {
        return Ok(CellValue::Error(
            "INDEX row_num/column_num out of bounds for array".to_string(),
        ));
    }

    let idx = ((row_num - 1) * cols + (col_num - 1)) as usize;
    Ok(array.values[idx].clone())
}

pub(crate) fn eval_match<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "MATCH expects 2 or 3 arguments (lookup_value, lookup_array, [match_type])".to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0])?;
    let lookup_array = flatten_range_expr(ctx, current_sheet, &args[1])?;

    if !is_1d(&lookup_array) {
        return Ok(CellValue::Error(
            "MATCH lookup_array must be a one-dimensional range".to_string(),
        ));
    }

    let match_type = if args.len() == 3 {
        let mt_val = evaluate_expression(ctx, current_sheet, &args[2])?;
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

pub(crate) fn eval_xmatch<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 4 {
        return Ok(CellValue::Error(
            "XMATCH expects 2 to 4 arguments (lookup_value, lookup_array, [match_mode], [search_mode])".to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0])?;
    let lookup_array = flatten_range_expr(ctx, current_sheet, &args[1])?;

    if !is_1d(&lookup_array) {
        return Ok(CellValue::Error(
            "XMATCH lookup_array must be a one-dimensional range".to_string(),
        ));
    }

    if args.len() >= 3 {
        let mm_val = evaluate_expression(ctx, current_sheet, &args[2])?;
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

pub(crate) fn eval_vlookup<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "VLOOKUP expects 3 or 4 arguments (lookup_value, table_array, col_index_num, [range_lookup])"
                .to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0])?;
    let table = flatten_range_expr(ctx, current_sheet, &args[1])?;

    let col_index_val = evaluate_expression(ctx, current_sheet, &args[2])?;
    let col_index = match to_number(&col_index_val) {
        Some(n) if n >= 1.0 => n as i64,
        _ => {
            return Ok(CellValue::Error(
                "VLOOKUP col_index_num must be a positive number".to_string(),
            ));
        },
    };

    let exact_match_only = if args.len() == 4 {
        let rl_val = evaluate_expression(ctx, current_sheet, &args[3])?;
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

pub(crate) fn eval_hlookup<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "HLOOKUP expects 3 or 4 arguments (lookup_value, table_array, row_index_num, [range_lookup])"
                .to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0])?;
    let table = flatten_range_expr(ctx, current_sheet, &args[1])?;

    let row_index_val = evaluate_expression(ctx, current_sheet, &args[2])?;
    let row_index = match to_number(&row_index_val) {
        Some(n) if n >= 1.0 => n as i64,
        _ => {
            return Ok(CellValue::Error(
                "HLOOKUP row_index_num must be a positive number".to_string(),
            ));
        },
    };

    let exact_match_only = if args.len() == 4 {
        let rl_val = evaluate_expression(ctx, current_sheet, &args[3])?;
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

    // First row contains the lookup keys.
    for c in 0..cols {
        let key = &table.values[c as usize];
        if values_equal(&lookup_val, key) {
            let idx = ((row_index - 1) * cols + c) as usize;
            return Ok(table.values[idx].clone());
        }
    }

    Ok(CellValue::Error("HLOOKUP: value not found".to_string()))
}

pub(crate) fn eval_xlookup<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 6 {
        return Ok(CellValue::Error(
            "XLOOKUP expects 3 to 6 arguments (lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])"
                .to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0])?;
    let lookup_array = flatten_range_expr(ctx, current_sheet, &args[1])?;
    let return_array = flatten_range_expr(ctx, current_sheet, &args[2])?;

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
        Some(evaluate_expression(ctx, current_sheet, &args[3])?)
    } else {
        None
    };

    if args.len() >= 5 {
        let mm_val = evaluate_expression(ctx, current_sheet, &args[4])?;
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

fn is_1d(range: &FlatRange) -> bool {
    range.rows == 1 || range.cols == 1
}

fn find_exact_match_index(lookup_val: &CellValue, values: &[CellValue]) -> Option<usize> {
    for (idx, v) in values.iter().enumerate() {
        if values_equal(lookup_val, v) {
            return Some(idx);
        }
    }
    None
}

fn values_equal(a: &CellValue, b: &CellValue) -> bool {
    match (to_number(a), to_number(b)) {
        (Some(x), Some(y)) => x == y,
        _ => to_text(a) == to_text(b),
    }
}
