use std::result::Result as StdResult;

use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, for_each_value_in_expr, to_number};

pub(crate) async fn eval_min(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut min_val: Option<f64> = None;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                min_val = Some(match min_val {
                    Some(cur) => cur.min(n),
                    None => n,
                });
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Float(min_val.unwrap_or(0.0)))
}

pub(crate) async fn eval_max(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut max_val: Option<f64> = None;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                max_val = Some(match max_val {
                    Some(cur) => cur.max(n),
                    None => n,
                });
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Float(max_val.unwrap_or(0.0)))
}

pub(crate) async fn eval_mina(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut min_val: Option<f64> = None;
    let mut error: Option<CellValue> = None;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if error.is_some() {
                return Ok(());
            }
            match coerce_for_mina_maxa(v) {
                Ok(Some(value)) => {
                    min_val = Some(match min_val {
                        Some(cur) => cur.min(value),
                        None => value,
                    });
                },
                Ok(None) => {},
                Err(msg) => {
                    error = Some(CellValue::Error(msg));
                },
            }
            Ok(())
        })
        .await?;
        if error.is_some() {
            break;
        }
    }

    if let Some(err) = error {
        return Ok(err);
    }

    match min_val {
        Some(value) => Ok(CellValue::Float(value)),
        None => Ok(CellValue::Float(0.0)),
    }
}

pub(crate) async fn eval_maxa(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut max_val: Option<f64> = None;
    let mut error: Option<CellValue> = None;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if error.is_some() {
                return Ok(());
            }
            match coerce_for_mina_maxa(v) {
                Ok(Some(value)) => {
                    max_val = Some(match max_val {
                        Some(cur) => cur.max(value),
                        None => value,
                    });
                },
                Ok(None) => {},
                Err(msg) => {
                    error = Some(CellValue::Error(msg));
                },
            }
            Ok(())
        })
        .await?;
        if error.is_some() {
            break;
        }
    }

    if let Some(err) = error {
        return Ok(err);
    }

    match max_val {
        Some(value) => Ok(CellValue::Float(value)),
        None => Ok(CellValue::Float(0.0)),
    }
}

fn coerce_for_mina_maxa(value: &CellValue) -> StdResult<Option<f64>, String> {
    match value {
        CellValue::Empty => Ok(None),
        CellValue::Bool(true) => Ok(Some(1.0)),
        CellValue::Bool(false) => Ok(Some(0.0)),
        CellValue::Int(i) => Ok(Some(*i as f64)),
        CellValue::Float(f) => Ok(Some(*f)),
        CellValue::DateTime(d) => Ok(Some(*d)),
        CellValue::String(_) => Ok(Some(0.0)),
        CellValue::Error(msg) => Err(msg.clone()),
        CellValue::Formula {
            cached_value: Some(inner),
            ..
        } => coerce_for_mina_maxa(inner),
        CellValue::Formula { .. } => Ok(None),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::engine::test_helpers::TestEngine;
    use crate::sheet::eval::parser::Expr;

    fn num_expr(n: f64) -> Expr {
        if n == n.floor() {
            Expr::Literal(CellValue::Int(n as i64))
        } else {
            Expr::Literal(CellValue::Float(n))
        }
    }

    #[tokio::test]
    async fn test_eval_min_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(10));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(5));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(20));
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_min(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 5.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_min_empty() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_min(ctx, "Sheet1", &args).await.unwrap();
        // MIN returns 0 when no numeric values found
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_max_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(10));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(5));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(20));
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_max(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 20.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_mina_includes_bool() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Bool(true)); // 1.0
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(0));
        engine.set_cell("Sheet1", 2, 0, CellValue::String("text".to_string())); // 0.0
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_mina(ctx, "Sheet1", &args).await.unwrap();
        // MINA includes text as 0, bool as 0/1 - min should be 0
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_maxa_includes_bool() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Bool(true)); // 1.0
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(0));
        engine.set_cell("Sheet1", 2, 0, CellValue::String("text".to_string())); // 0.0
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_maxa(ctx, "Sheet1", &args).await.unwrap();
        // MAXA includes text as 0, bool as 0/1 - max should be 1
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }
}
