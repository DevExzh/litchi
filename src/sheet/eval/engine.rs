//! Runtime evaluation of parsed formula expressions.
//!
//! This module operates on the small expression AST defined in
//! `sheet::eval::parser` and evaluates it against an evaluation engine.
//!
//! The initial implementation is intentionally conservative and supports
//! only scalar arithmetic over numeric literals and single-cell references.

use crate::sheet::{CellValue, Result};

pub(crate) use super::EngineCtx;
use super::parser::Expr;
pub(crate) use registry::DispatchCtx;

pub(crate) type EvalCtx<'a> = &'a dyn DispatchCtx;

#[derive(Debug, Clone)]
pub(crate) enum ResolvedName {
    Cell { sheet: String, row: u32, col: u32 },
    Range(super::parser::RangeRef),
}
pub(crate) trait ReferenceResolver {
    fn resolve_name(&self, _current_sheet: &str, _name: &str) -> Result<Option<ResolvedName>> {
        Ok(None)
    }
}

mod aggregate;
mod bin_op;
mod criteria;
mod criteria_aggs;
mod database;
mod date_time;
mod dispatch;
mod engineering;
mod financial;
mod info;
mod logical;
mod lookup;
mod math;
mod registry;
mod statistical;
mod text;
mod web;

#[cfg(test)]
mod tests;

/// Evaluate a parsed expression in the context of an evaluation engine.
pub(crate) async fn evaluate_expression(
    ctx: &dyn DispatchCtx,
    current_sheet: &str,
    expr: &Expr,
) -> Result<CellValue> {
    let value = match expr {
        Expr::Literal(v) => v.clone(),
        Expr::Reference { sheet, row, col } => {
            let sheet_name = sheet.as_str();
            ctx.get_cell_value(sheet_name, *row, *col).await?
        },
        Expr::Name(name) => match ctx.resolve_name(current_sheet, name.as_str())? {
            Some(ResolvedName::Cell { sheet, row, col }) => {
                ctx.get_cell_value(sheet.as_str(), row, col).await?
            },
            Some(ResolvedName::Range(_range)) => CellValue::Error(format!(
                "Named range '{}' cannot be used as a scalar expression",
                name
            )),
            None => CellValue::Error(format!("Unknown name: {}", name)),
        },
        Expr::Range(_) => {
            // A bare range used as a scalar expression is not supported in this
            // MVP engine. Ranges are currently only meaningful as arguments to
            // aggregate functions like SUM.
            CellValue::Error("Range cannot be used as a scalar expression".to_string())
        },
        Expr::UnaryMinus(inner) => {
            let v = Box::pin(evaluate_expression(ctx, current_sheet, inner)).await?;
            match v {
                CellValue::Int(i) => CellValue::Int(-i),
                CellValue::Float(f) => CellValue::Float(-f),
                other => CellValue::Error(format!("Unary minus on non-numeric value: {:?}", other)),
            }
        },
        Expr::Binary { op, left, right } => {
            let left_val = Box::pin(evaluate_expression(ctx, current_sheet, left)).await?;
            let right_val = Box::pin(evaluate_expression(ctx, current_sheet, right)).await?;
            bin_op::eval_binary_op(*op, left_val, right_val)
        },
        Expr::FunctionCall { name, args } => {
            dispatch::eval_function(ctx, current_sheet, name, args).await?
        },
    };

    Ok(value)
}

pub(crate) fn to_number(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Int(i) => Some(*i as f64),
        CellValue::Float(f) => Some(*f),
        CellValue::DateTime(d) => Some(*d),
        _ => None,
    }
}

pub(crate) fn to_bool(value: &CellValue) -> bool {
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

pub(crate) fn to_text(value: &CellValue) -> String {
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

pub(crate) fn is_blank(value: &CellValue) -> bool {
    match value {
        CellValue::Empty => true,
        CellValue::String(s) => s.is_empty(),
        _ => false,
    }
}

pub(crate) async fn for_each_value_in_expr<F>(
    ctx: &dyn DispatchCtx,
    current_sheet: &str,
    expr: &Expr,
    mut f: F,
) -> Result<()>
where
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
                    let v = ctx.get_cell_value(range.sheet.as_str(), row, col).await?;
                    f(&v)?;
                }
            }
        },
        Expr::Name(name) => match ctx.resolve_name(current_sheet, name.as_str())? {
            Some(ResolvedName::Cell { sheet, row, col }) => {
                let v = ctx.get_cell_value(sheet.as_str(), row, col).await?;
                f(&v)?;
            },
            Some(ResolvedName::Range(range)) => {
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
                        let v = ctx.get_cell_value(range.sheet.as_str(), row, col).await?;
                        f(&v)?;
                    }
                }
            },
            None => {
                let v = CellValue::Error(format!("Unknown name: {}", name));
                f(&v)?;
            },
        },
        other => {
            let v = evaluate_expression(ctx, current_sheet, other).await?;
            f(&v)?;
        },
    }
    Ok(())
}

#[derive(Clone)]
pub(crate) struct FlatRange {
    pub(crate) values: Vec<CellValue>,
    pub(crate) rows: usize,
    pub(crate) cols: usize,
}

pub(crate) async fn flatten_range_expr(
    ctx: &dyn DispatchCtx,
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
                    let v = ctx.get_cell_value(range.sheet.as_str(), row, col).await?;
                    values.push(v);
                }
            }

            Ok(FlatRange { values, rows, cols })
        },
        Expr::Name(name) => match ctx.resolve_name(current_sheet, name.as_str())? {
            Some(ResolvedName::Cell { sheet, row, col }) => {
                let v = ctx.get_cell_value(sheet.as_str(), row, col).await?;
                Ok(FlatRange {
                    values: vec![v],
                    rows: 1,
                    cols: 1,
                })
            },
            Some(ResolvedName::Range(range)) => {
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
                        let v = ctx.get_cell_value(range.sheet.as_str(), row, col).await?;
                        values.push(v);
                    }
                }

                Ok(FlatRange { values, rows, cols })
            },
            None => {
                let v = CellValue::Error(format!("Unknown name: {}", name));
                Ok(FlatRange {
                    values: vec![v],
                    rows: 1,
                    cols: 1,
                })
            },
        },
        other => {
            let v = evaluate_expression(ctx, current_sheet, other).await?;
            Ok(FlatRange {
                values: vec![v],
                rows: 1,
                cols: 1,
            })
        },
    }
}
