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

#[cfg(test)]
pub(crate) mod test_helpers {
    use super::*;
    use crate::sheet::Result;
    use crate::sheet::eval::BoxFuture;
    use std::collections::HashMap;
    use std::sync::{Arc, RwLock};

    /// A simple test engine for creating EvalCtx in tests.
    pub(crate) struct TestEngine {
        cells: Arc<RwLock<HashMap<(String, u32, u32), CellValue>>>,
        current_pos: Arc<RwLock<Option<(String, u32, u32)>>>,
        sheet_count: usize,
    }

    impl TestEngine {
        pub(crate) fn new() -> Self {
            Self {
                cells: Arc::new(RwLock::new(HashMap::new())),
                current_pos: Arc::new(RwLock::new(None)),
                sheet_count: 1,
            }
        }

        /// Returns a reference to self as an EvalCtx
        pub(crate) fn ctx(&self) -> EvalCtx<'_> {
            self
        }

        /// Set a cell value in the test engine
        pub(crate) fn set_cell(&self, sheet: &str, row: u32, col: u32, value: CellValue) {
            let mut cells = self.cells.write().unwrap();
            cells.insert((sheet.to_string(), row, col), value);
        }

        /// Add a range of values starting at the given position
        pub(crate) fn add_range(
            &self,
            sheet: &str,
            start_row: u32,
            start_col: u32,
            rows: usize,
            cols: usize,
            values: Vec<CellValue>,
        ) {
            let mut cells = self.cells.write().unwrap();
            for (idx, value) in values.iter().enumerate() {
                let r = idx / cols;
                let c = idx % cols;
                if r < rows {
                    cells.insert(
                        (
                            sheet.to_string(),
                            start_row + r as u32,
                            start_col + c as u32,
                        ),
                        value.clone(),
                    );
                }
            }
        }

        /// Set the current position for ROW/COLUMN functions
        pub(crate) fn set_current_position(&self, sheet: &str, row: u32, col: u32) {
            let mut pos = self.current_pos.write().unwrap();
            *pos = Some((sheet.to_string(), row, col));
        }
    }

    impl EngineCtx for TestEngine {
        fn get_cell_value<'a>(
            &'a self,
            sheet_name: &'a str,
            row: u32,
            col: u32,
        ) -> BoxFuture<'a, Result<CellValue>> {
            let cells = self.cells.clone();
            let sheet = sheet_name.to_string();
            Box::pin(async move {
                let cells = cells.read().unwrap();
                Ok(cells
                    .get(&(sheet, row, col))
                    .cloned()
                    .unwrap_or(CellValue::Empty))
            })
        }

        fn current_position(&self) -> Option<(String, u32, u32)> {
            self.current_pos.read().unwrap().clone()
        }

        fn raw_cell_value<'a>(
            &'a self,
            sheet_name: &'a str,
            row: u32,
            col: u32,
        ) -> BoxFuture<'a, Result<CellValue>> {
            self.get_cell_value(sheet_name, row, col)
        }

        fn is_1904_date_system(&self) -> bool {
            false
        }

        #[cfg(feature = "eval_engine_web_functions")]
        fn http_client(&self) -> &reqwest::Client {
            panic!("TestEngine does not support HTTP client")
        }

        fn get_sheet_index(&self, _name: &str) -> Option<usize> {
            Some(0)
        }

        fn get_sheet_count(&self) -> usize {
            self.sheet_count
        }
    }

    impl ReferenceResolver for TestEngine {}
}

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
