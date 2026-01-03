use crate::sheet::{CellValue, Result};

use super::super::parser::Expr;
use super::criteria::{Criteria, matches_criteria, parse_criteria};
use super::{EvalCtx, FlatRange, flatten_range_expr, to_number, to_text};

#[derive(Clone, Copy)]
enum Extremum {
    Min,
    Max,
}

async fn minmax_ifs(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    func_name: &str,
    extremum: Extremum,
) -> Result<CellValue> {
    if args.len() < 3 || args.len().is_multiple_of(2) {
        return Ok(CellValue::Error(format!(
            "{func_name} expects 3 or more arguments ({func_name_lower} range, criteria_range1, criteria1, ...)",
            func_name_lower = func_name.to_lowercase()
        )));
    }

    let target_range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;

    let mut crit_ranges: Vec<FlatRange> = Vec::new();
    let mut crits: Vec<Criteria> = Vec::new();

    let mut i = 1;
    while i + 1 < args.len() {
        let r = flatten_range_expr(ctx, current_sheet, &args[i]).await?;
        if r.rows != target_range.rows || r.cols != target_range.cols {
            return Ok(CellValue::Error(format!(
                "{func_name} criteria ranges must have the same size as the target range"
            )));
        }
        let crit_val = super::evaluate_expression(ctx, current_sheet, &args[i + 1]).await?;
        let crit_str = to_text(&crit_val);
        let c = match parse_criteria(&crit_str) {
            Some(c) => c,
            None => {
                return Ok(CellValue::Error(format!(
                    "{func_name} received an invalid criteria"
                )));
            },
        };
        crit_ranges.push(r);
        crits.push(c);
        i += 2;
    }

    let mut best: Option<f64> = None;
    for idx in 0..target_range.values.len() {
        let mut ok = true;
        for (r, c) in crit_ranges.iter().zip(crits.iter()) {
            if !matches_criteria(&r.values[idx], c) {
                ok = false;
                break;
            }
        }
        if ok && let Some(value) = to_number(&target_range.values[idx]) {
            best = Some(match (best, extremum) {
                (Some(current), Extremum::Min) => current.min(value),
                (Some(current), Extremum::Max) => current.max(value),
                (None, _) => value,
            });
        }
    }

    Ok(CellValue::Float(best.unwrap_or(0.0)))
}

pub(crate) async fn eval_minifs(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    minmax_ifs(ctx, current_sheet, args, "MINIFS", Extremum::Min).await
}

pub(crate) async fn eval_maxifs(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    minmax_ifs(ctx, current_sheet, args, "MAXIFS", Extremum::Max).await
}

pub(crate) async fn eval_sumif(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "SUMIF expects 2 or 3 arguments (range, criteria, [sum_range])".to_string(),
        ));
    }

    let range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let crit_val = super::evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let crit_str = to_text(&crit_val);
    let criteria = match parse_criteria(&crit_str) {
        Some(c) => c,
        None => {
            return Ok(CellValue::Error("Invalid SUMIF criteria".to_string()));
        },
    };

    let sum_range = if args.len() == 3 {
        let sr = flatten_range_expr(ctx, current_sheet, &args[2]).await?;
        if sr.rows != range.rows || sr.cols != range.cols {
            return Ok(CellValue::Error(
                "SUMIF range and sum_range must have the same size".to_string(),
            ));
        }
        sr
    } else {
        FlatRange {
            values: range.values.clone(),
            rows: range.rows,
            cols: range.cols,
        }
    };

    let mut total = 0.0f64;
    for i in 0..range.values.len() {
        if matches_criteria(&range.values[i], &criteria)
            && let Some(n) = to_number(&sum_range.values[i])
        {
            total += n;
        }
    }

    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_countif(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "COUNTIF expects 2 arguments (range, criteria)".to_string(),
        ));
    }

    let range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let crit_val = super::evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let crit_str = to_text(&crit_val);
    let criteria = match parse_criteria(&crit_str) {
        Some(c) => c,
        None => {
            return Ok(CellValue::Error("Invalid COUNTIF criteria".to_string()));
        },
    };

    let mut count = 0i64;
    for v in &range.values {
        if matches_criteria(v, &criteria) {
            count += 1;
        }
    }

    Ok(CellValue::Int(count))
}

pub(crate) async fn eval_averageif(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "AVERAGEIF expects 2 or 3 arguments (range, criteria, [average_range])".to_string(),
        ));
    }

    let range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let crit_val = super::evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let crit_str = to_text(&crit_val);
    let criteria = match parse_criteria(&crit_str) {
        Some(c) => c,
        None => {
            return Ok(CellValue::Error("Invalid AVERAGEIF criteria".to_string()));
        },
    };

    let avg_range = if args.len() == 3 {
        let ar = flatten_range_expr(ctx, current_sheet, &args[2]).await?;
        if ar.rows != range.rows || ar.cols != range.cols {
            return Ok(CellValue::Error(
                "AVERAGEIF range and average_range must have the same size".to_string(),
            ));
        }
        ar
    } else {
        FlatRange {
            values: range.values.clone(),
            rows: range.rows,
            cols: range.cols,
        }
    };

    let mut total = 0.0f64;
    let mut count = 0u64;

    for i in 0..range.values.len() {
        if matches_criteria(&range.values[i], &criteria)
            && let Some(n) = to_number(&avg_range.values[i])
        {
            total += n;
            count += 1;
        }
    }

    if count == 0 {
        return Ok(CellValue::Error(
            "AVERAGEIF has no matching numeric values".to_string(),
        ));
    }

    Ok(CellValue::Float(total / count as f64))
}

pub(crate) async fn eval_sumifs(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len().is_multiple_of(2) {
        return Ok(CellValue::Error(
            "SUMIFS expects 3 or more arguments (sum_range, criteria_range1, criteria1, ...)"
                .to_string(),
        ));
    }

    let sum_range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;

    let mut crit_ranges: Vec<FlatRange> = Vec::new();
    let mut crits: Vec<Criteria> = Vec::new();

    let mut i = 1;
    while i + 1 < args.len() {
        let r = flatten_range_expr(ctx, current_sheet, &args[i]).await?;
        if r.rows != sum_range.rows || r.cols != sum_range.cols {
            return Ok(CellValue::Error(
                "SUMIFS criteria ranges must have the same size as sum_range".to_string(),
            ));
        }
        let crit_val = super::evaluate_expression(ctx, current_sheet, &args[i + 1]).await?;
        let crit_str = to_text(&crit_val);
        let c = match parse_criteria(&crit_str) {
            Some(c) => c,
            None => {
                return Ok(CellValue::Error("Invalid SUMIFS criteria".to_string()));
            },
        };
        crit_ranges.push(r);
        crits.push(c);
        i += 2;
    }

    let mut total = 0.0f64;
    for idx in 0..sum_range.values.len() {
        let mut ok = true;
        for (r, c) in crit_ranges.iter().zip(crits.iter()) {
            if !matches_criteria(&r.values[idx], c) {
                ok = false;
                break;
            }
        }
        if ok && let Some(n) = to_number(&sum_range.values[idx]) {
            total += n;
        }
    }

    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_countifs(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || !args.len().is_multiple_of(2) {
        return Ok(CellValue::Error(
            "COUNTIFS expects an even number of arguments (criteria_range1, criteria1, ...)"
                .to_string(),
        ));
    }

    let first_range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let mut crit_ranges: Vec<FlatRange> = Vec::new();
    let mut crits: Vec<Criteria> = Vec::new();

    let mut i = 0;
    while i + 1 < args.len() {
        let r = if i == 0 {
            first_range.clone()
        } else {
            flatten_range_expr(ctx, current_sheet, &args[i]).await?
        };
        if r.rows != first_range.rows || r.cols != first_range.cols {
            return Ok(CellValue::Error(
                "COUNTIFS criteria ranges must all have the same size".to_string(),
            ));
        }
        let crit_val = super::evaluate_expression(ctx, current_sheet, &args[i + 1]).await?;
        let crit_str = to_text(&crit_val);
        let c = match parse_criteria(&crit_str) {
            Some(c) => c,
            None => {
                return Ok(CellValue::Error("Invalid COUNTIFS criteria".to_string()));
            },
        };
        crit_ranges.push(r);
        crits.push(c);
        i += 2;
    }

    let mut count = 0i64;
    for idx in 0..first_range.values.len() {
        let mut ok = true;
        for (r, c) in crit_ranges.iter().zip(crits.iter()) {
            if !matches_criteria(&r.values[idx], c) {
                ok = false;
                break;
            }
        }
        if ok {
            count += 1;
        }
    }

    Ok(CellValue::Int(count))
}

pub(crate) async fn eval_averageifs(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len().is_multiple_of(2) {
        return Ok(CellValue::Error(
            "AVERAGEIFS expects 3 or more arguments (average_range, criteria_range1, criteria1, ...)".to_string(),
        ));
    }

    let avg_range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let mut crit_ranges: Vec<FlatRange> = Vec::new();
    let mut crits: Vec<Criteria> = Vec::new();

    let mut i = 1;
    while i + 1 < args.len() {
        let r = flatten_range_expr(ctx, current_sheet, &args[i]).await?;
        if r.rows != avg_range.rows || r.cols != avg_range.cols {
            return Ok(CellValue::Error(
                "AVERAGEIFS criteria ranges must have the same size as average_range".to_string(),
            ));
        }
        let crit_val = super::evaluate_expression(ctx, current_sheet, &args[i + 1]).await?;
        let crit_str = to_text(&crit_val);
        let c = match parse_criteria(&crit_str) {
            Some(c) => c,
            None => {
                return Ok(CellValue::Error("Invalid AVERAGEIFS criteria".to_string()));
            },
        };
        crit_ranges.push(r);
        crits.push(c);
        i += 2;
    }

    let mut total = 0.0f64;
    let mut count = 0u64;

    for idx in 0..avg_range.values.len() {
        let mut ok = true;
        for (r, c) in crit_ranges.iter().zip(crits.iter()) {
            if !matches_criteria(&r.values[idx], c) {
                ok = false;
                break;
            }
        }
        if ok && let Some(n) = to_number(&avg_range.values[idx]) {
            total += n;
            count += 1;
        }
    }

    if count == 0 {
        return Ok(CellValue::Error(
            "AVERAGEIFS has no matching numeric values".to_string(),
        ));
    }

    Ok(CellValue::Float(total / count as f64))
}
