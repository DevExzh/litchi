//! Runtime evaluation of parsed formula expressions.
//!
//! This module operates on the small expression AST defined in
//! `sheet::eval::parser` and evaluates it against an evaluation engine.
//!
//! The initial implementation is intentionally conservative and supports
//! only scalar arithmetic over numeric literals and single-cell references.

use crate::sheet::{CellValue, Result};

use std::f64::consts::PI;

pub(crate) use super::EngineCtx;
use super::parser::Expr;

mod aggregate;
mod bin_op;
mod criteria;
mod criteria_aggs;
mod date_time;
mod financial;
mod logical;
mod lookup;
mod math;
mod statistical;
mod text;

/// Evaluate a parsed expression in the context of an evaluation engine.
pub(crate) fn evaluate_expression<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    expr: &Expr,
) -> Result<CellValue> {
    let value = match expr {
        Expr::Literal(v) => v.clone(),
        Expr::Reference { sheet, row, col } => {
            let sheet_name = sheet.as_str();
            ctx.get_cell_value(sheet_name, *row, *col)?
        },
        Expr::Range(_) => {
            // A bare range used as a scalar expression is not supported in this
            // MVP engine. Ranges are currently only meaningful as arguments to
            // aggregate functions like SUM.
            CellValue::Error("Range cannot be used as a scalar expression".to_string())
        },
        Expr::UnaryMinus(inner) => {
            let v = evaluate_expression(ctx, current_sheet, inner)?;
            match v {
                CellValue::Int(i) => CellValue::Int(-i),
                CellValue::Float(f) => CellValue::Float(-f),
                other => CellValue::Error(format!("Unary minus on non-numeric value: {:?}", other)),
            }
        },
        Expr::Binary { op, left, right } => {
            let left_val = evaluate_expression(ctx, current_sheet, left)?;
            let right_val = evaluate_expression(ctx, current_sheet, right)?;
            bin_op::eval_binary_op(*op, left_val, right_val)
        },
        Expr::FunctionCall { name, args } => eval_function(ctx, current_sheet, name, args)?,
    };

    Ok(value)
}

fn to_number(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Int(i) => Some(*i as f64),
        CellValue::Float(f) => Some(*f),
        CellValue::DateTime(d) => Some(*d),
        _ => None,
    }
}

fn to_bool(value: &CellValue) -> bool {
    match value {
        CellValue::Bool(b) => *b,
        CellValue::Int(i) => *i != 0,
        CellValue::Float(f) | CellValue::DateTime(f) => *f != 0.0,
        CellValue::String(s) => !s.is_empty(),
        CellValue::Empty => false,
        CellValue::Error(_) => false,
        CellValue::Formula { .. } => false,
    }
}

fn to_text(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Bool(true) => "TRUE".to_string(),
        CellValue::Bool(false) => "FALSE".to_string(),
        CellValue::Int(i) => i.to_string(),
        CellValue::Float(f) => f.to_string(),
        CellValue::DateTime(d) => d.to_string(),
        CellValue::String(s) => s.clone(),
        CellValue::Error(e) => e.clone(),
        CellValue::Formula { .. } => "[FORMULA]".to_string(),
    }
}

fn is_blank(value: &CellValue) -> bool {
    match value {
        CellValue::Empty => true,
        CellValue::String(s) => s.is_empty(),
        _ => false,
    }
}

fn for_each_value_in_expr<C, F>(ctx: &C, current_sheet: &str, expr: &Expr, mut f: F) -> Result<()>
where
    C: EngineCtx + ?Sized,
    F: FnMut(&CellValue) -> Result<()>,
{
    match expr {
        Expr::Range(range) => {
            let (sr, er) = if range.start_row <= range.end_row {
                (range.start_row, range.end_row)
            } else {
                (range.end_row, range.start_row)
            };
            let (sc, ec) = if range.start_col <= range.end_col {
                (range.start_col, range.end_col)
            } else {
                (range.end_col, range.start_col)
            };

            for row in sr..=er {
                for col in sc..=ec {
                    let v = ctx.get_cell_value(range.sheet.as_str(), row, col)?;
                    f(&v)?;
                }
            }
        },
        other => {
            let v = evaluate_expression(ctx, current_sheet, other)?;
            f(&v)?;
        },
    }
    Ok(())
}

#[derive(Clone)]
struct FlatRange {
    values: Vec<CellValue>,
    rows: usize,
    cols: usize,
}

fn flatten_range_expr<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    expr: &Expr,
) -> Result<FlatRange> {
    match expr {
        Expr::Range(range) => {
            let (sr, er) = if range.start_row <= range.end_row {
                (range.start_row, range.end_row)
            } else {
                (range.end_row, range.start_row)
            };
            let (sc, ec) = if range.start_col <= range.end_col {
                (range.start_col, range.end_col)
            } else {
                (range.end_col, range.start_col)
            };

            let rows = (er - sr + 1) as usize;
            let cols = (ec - sc + 1) as usize;
            let mut values = Vec::with_capacity(rows * cols);

            for row in sr..=er {
                for col in sc..=ec {
                    let v = ctx.get_cell_value(range.sheet.as_str(), row, col)?;
                    values.push(v);
                }
            }

            Ok(FlatRange { values, rows, cols })
        },
        other => {
            let v = evaluate_expression(ctx, current_sheet, other)?;
            Ok(FlatRange {
                values: vec![v],
                rows: 1,
                cols: 1,
            })
        },
    }
}

fn eval_function<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    name: &str,
    args: &[Expr],
) -> Result<CellValue> {
    match name {
        // Basic SUM implementation: sums numeric values from scalar
        // expressions and ranges. Non-numeric values are ignored.
        "SUM" => aggregate::eval_sum(ctx, current_sheet, args),
        "PRODUCT" => aggregate::eval_product(ctx, current_sheet, args),
        "MIN" => aggregate::eval_min(ctx, current_sheet, args),
        "MAX" => aggregate::eval_max(ctx, current_sheet, args),
        "AVERAGE" => aggregate::eval_average(ctx, current_sheet, args),
        "COUNT" => aggregate::eval_count(ctx, current_sheet, args),
        "COUNTA" => aggregate::eval_counta(ctx, current_sheet, args),
        "SUMIF" => criteria_aggs::eval_sumif(ctx, current_sheet, args),
        "SUMIFS" => criteria_aggs::eval_sumifs(ctx, current_sheet, args),
        "COUNTIF" => criteria_aggs::eval_countif(ctx, current_sheet, args),
        "COUNTIFS" => criteria_aggs::eval_countifs(ctx, current_sheet, args),
        "AVERAGEIF" => criteria_aggs::eval_averageif(ctx, current_sheet, args),
        "AVERAGEIFS" => criteria_aggs::eval_averageifs(ctx, current_sheet, args),

        "ABS" => math::eval_abs(ctx, current_sheet, args),
        "INT" => math::eval_int(ctx, current_sheet, args),
        "ROUND" => math::eval_round(ctx, current_sheet, args),
        "ROUNDDOWN" => math::eval_rounddown(ctx, current_sheet, args),
        "ROUNDUP" => math::eval_roundup(ctx, current_sheet, args),
        "FLOOR" => math::eval_floor(ctx, current_sheet, args),
        "CEILING" => math::eval_ceiling(ctx, current_sheet, args),
        "POWER" => math::eval_power(ctx, current_sheet, args),
        "SQRT" => math::eval_sqrt(ctx, current_sheet, args),
        "EXP" => math::eval_exp(ctx, current_sheet, args),
        "LN" => math::eval_ln(ctx, current_sheet, args),
        "LOG10" => math::eval_log10(ctx, current_sheet, args),
        "PI" => Ok(CellValue::Float(PI)),
        "SIN" => math::eval_sin(ctx, current_sheet, args),
        "COS" => math::eval_cos(ctx, current_sheet, args),
        "TAN" => math::eval_tan(ctx, current_sheet, args),
        "ASIN" => math::eval_asin(ctx, current_sheet, args),
        "ACOS" => math::eval_acos(ctx, current_sheet, args),
        "ATAN" => math::eval_atan(ctx, current_sheet, args),
        "ATAN2" => math::eval_atan2(ctx, current_sheet, args),

        "IF" => logical::eval_if(ctx, current_sheet, args),
        "AND" => logical::eval_and(ctx, current_sheet, args),
        "OR" => logical::eval_or(ctx, current_sheet, args),
        "NOT" => logical::eval_not(ctx, current_sheet, args),

        "LEN" => text::eval_len(ctx, current_sheet, args),
        "LOWER" => text::eval_lower(ctx, current_sheet, args),
        "UPPER" => text::eval_upper(ctx, current_sheet, args),
        "TRIM" => text::eval_trim(ctx, current_sheet, args),
        "CONCAT" | "CONCATENATE" => text::eval_concat(ctx, current_sheet, args),
        "TEXTJOIN" => text::eval_textjoin(ctx, current_sheet, args),

        "TODAY" => date_time::eval_today(ctx, current_sheet, args),
        "NOW" => date_time::eval_now(ctx, current_sheet, args),
        "DATE" => date_time::eval_date(ctx, current_sheet, args),
        "TIME" => date_time::eval_time(ctx, current_sheet, args),
        "DATEVALUE" => date_time::eval_datevalue(ctx, current_sheet, args),
        "TIMEVALUE" => date_time::eval_timevalue(ctx, current_sheet, args),
        "EDATE" => date_time::eval_edate(ctx, current_sheet, args),
        "EOMONTH" => date_time::eval_eomonth(ctx, current_sheet, args),
        "WORKDAY" => date_time::eval_workday(ctx, current_sheet, args),
        "WORKDAY.INTL" => date_time::eval_workday_intl(ctx, current_sheet, args),
        "NETWORKDAYS" => date_time::eval_networkdays(ctx, current_sheet, args),

        "INDEX" => lookup::eval_index(ctx, current_sheet, args),
        "MATCH" => lookup::eval_match(ctx, current_sheet, args),
        "XMATCH" => lookup::eval_xmatch(ctx, current_sheet, args),
        "VLOOKUP" => lookup::eval_vlookup(ctx, current_sheet, args),
        "HLOOKUP" => lookup::eval_hlookup(ctx, current_sheet, args),
        "XLOOKUP" => lookup::eval_xlookup(ctx, current_sheet, args),
        "NORM.DIST" => statistical::eval_norm_dist(ctx, current_sheet, args),
        "NORM.S.INV" => statistical::eval_norm_s_inv(ctx, current_sheet, args),
        "PROB" => statistical::eval_prob(ctx, current_sheet, args),
        "CHISQ.DIST" => statistical::eval_chisq_dist(ctx, current_sheet, args),
        "CHISQ.DIST.RT" => statistical::eval_chisq_dist_rt(ctx, current_sheet, args),
        "CHISQ.INV" => statistical::eval_chisq_inv(ctx, current_sheet, args),
        "CHISQ.INV.RT" => statistical::eval_chisq_inv_rt(ctx, current_sheet, args),
        "T.DIST" => statistical::eval_t_dist(ctx, current_sheet, args),
        "T.DIST.2T" => statistical::eval_t_dist_2t(ctx, current_sheet, args),
        "T.DIST.RT" => statistical::eval_t_dist_rt(ctx, current_sheet, args),
        "T.INV" => statistical::eval_t_inv(ctx, current_sheet, args),
        "T.INV.2T" => statistical::eval_t_inv_2t(ctx, current_sheet, args),
        "F.DIST" => statistical::eval_f_dist(ctx, current_sheet, args),
        "F.DIST.RT" => statistical::eval_f_dist_rt(ctx, current_sheet, args),
        "F.INV" => statistical::eval_f_inv(ctx, current_sheet, args),
        "F.INV.RT" => statistical::eval_f_inv_rt(ctx, current_sheet, args),
        "PV" => financial::eval_pv(ctx, current_sheet, args),
        "FV" => financial::eval_fv(ctx, current_sheet, args),
        "RATE" => financial::eval_rate(ctx, current_sheet, args),
        "NPV" => financial::eval_npv(ctx, current_sheet, args),
        "IRR" => financial::eval_irr(ctx, current_sheet, args),
        "XNPV" => financial::eval_xnpv(ctx, current_sheet, args),
        "XIRR" => financial::eval_xirr(ctx, current_sheet, args),
        "YIELD" => financial::eval_yield(ctx, current_sheet, args),
        "DURATION" => financial::eval_duration(ctx, current_sheet, args),
        other => Ok(CellValue::Error(format!("Unsupported function: {}", other))),
    }
}
