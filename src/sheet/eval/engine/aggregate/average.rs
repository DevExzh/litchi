use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, for_each_value_in_expr, to_number};

pub(crate) async fn eval_average(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut total = 0.0f64;
    let mut count = 0u64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                total += n;
                count += 1;
            }
            Ok(())
        })
        .await?;
    }
    if count == 0 {
        return Ok(CellValue::Error("AVERAGE of empty set".to_string()));
    }
    Ok(CellValue::Float(total / count as f64))
}

pub(crate) async fn eval_averagea(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut total = 0.0f64;
    let mut count = 0u64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            match v {
                CellValue::Empty => {},
                CellValue::Bool(true) => {
                    total += 1.0;
                    count += 1;
                },
                CellValue::Bool(false) => {
                    count += 1;
                },
                CellValue::Int(i) => {
                    total += *i as f64;
                    count += 1;
                },
                CellValue::Float(f) => {
                    total += *f;
                    count += 1;
                },
                CellValue::DateTime(d) => {
                    total += *d;
                    count += 1;
                },
                CellValue::String(_) => {
                    count += 1;
                },
                CellValue::Error(_) => {},
                CellValue::Formula {
                    cached_value: Some(value),
                    ..
                } => match &**value {
                    CellValue::Bool(true) => {
                        total += 1.0;
                        count += 1;
                    },
                    CellValue::Bool(false) => {
                        count += 1;
                    },
                    CellValue::Int(i) => {
                        total += *i as f64;
                        count += 1;
                    },
                    CellValue::Float(f) => {
                        total += *f;
                        count += 1;
                    },
                    CellValue::DateTime(d) => {
                        total += *d;
                        count += 1;
                    },
                    CellValue::String(_) => {
                        count += 1;
                    },
                    _ => {},
                },
                CellValue::Formula { .. } => {
                    count += 1;
                },
            }
            Ok(())
        })
        .await?;
    }
    if count == 0 {
        return Ok(CellValue::Error("AVERAGEA of empty set".to_string()));
    }
    Ok(CellValue::Float(total / count as f64))
}

pub(crate) async fn eval_avedev(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut values = Vec::new();
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                values.push(n);
            }
            Ok(())
        })
        .await?;
    }

    if values.is_empty() {
        return Ok(CellValue::Error(
            "AVEDEV requires at least one numeric value".to_string(),
        ));
    }

    let mean = values.iter().sum::<f64>() / values.len() as f64;
    let total_dev = values.iter().map(|v| (v - mean).abs()).sum::<f64>();
    Ok(CellValue::Float(total_dev / values.len() as f64))
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
    async fn test_eval_average_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(10));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(20));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(30));
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_average(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 20.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_average_empty() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_average(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }

    #[tokio::test]
    async fn test_eval_average_ignores_non_numeric() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(10));
        engine.set_cell("Sheet1", 1, 0, CellValue::String("text".to_string()));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(30));
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_average(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 20.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_averagea_includes_bool_and_text() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(10));
        engine.set_cell("Sheet1", 1, 0, CellValue::Bool(true));
        engine.set_cell("Sheet1", 2, 0, CellValue::String("text".to_string()));
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_averagea(ctx, "Sheet1", &args).await.unwrap();
        // (10 + 1 + 0) / 3 = 3.666...
        match result {
            CellValue::Float(v) => assert!((v - 3.6666667).abs() < 0.001),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_averagea_bool_false() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Bool(false));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(10));
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 1,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_averagea(ctx, "Sheet1", &args).await.unwrap();
        // (0 + 10) / 2 = 5
        match result {
            CellValue::Float(v) => assert!((v - 5.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_avedev() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(4));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(5));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(6));
        engine.set_cell("Sheet1", 3, 0, CellValue::Int(7));
        engine.set_cell("Sheet1", 4, 0, CellValue::Int(5));
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 4,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_avedev(ctx, "Sheet1", &args).await.unwrap();
        // Mean = 5.4, AVEDEV = (|4-5.4| + |5-5.4| + |6-5.4| + |7-5.4| + |5-5.4|) / 5 = 0.88
        match result {
            CellValue::Float(v) => assert!((v - 0.88).abs() < 0.01),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_avedev_empty() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_avedev(ctx, "Sheet1", &args).await.unwrap();
        assert!(matches!(result, CellValue::Error(_)));
    }
}
