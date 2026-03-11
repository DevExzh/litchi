use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, flatten_range_expr, for_each_value_in_expr, to_number};

pub(crate) async fn eval_sum(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut total = 0.0f64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                total += n;
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_product(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut product = 1.0f64;
    let mut found_numeric = false;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                product *= n;
                found_numeric = true;
            }
            Ok(())
        })
        .await?;
    }
    if !found_numeric {
        Ok(CellValue::Float(0.0))
    } else {
        Ok(CellValue::Float(product))
    }
}

pub(crate) async fn eval_sumproduct(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "SUMPRODUCT expects at least 1 argument".to_string(),
        ));
    }

    let mut flattened_args = Vec::with_capacity(args.len());
    for arg in args {
        flattened_args.push(flatten_range_expr(ctx, current_sheet, arg).await?);
    }

    // All arrays must have the same dimensions
    let rows = flattened_args[0].rows;
    let cols = flattened_args[0].cols;
    for arg in &flattened_args[1..] {
        if arg.rows != rows || arg.cols != cols {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }
    }

    let mut total = 0.0;
    for i in 0..(rows * cols) {
        let mut product = 1.0;
        for arg in &flattened_args {
            let val = &arg.values[i];
            let n = to_number(val).unwrap_or(0.0);
            product *= n;
        }
        total += product;
    }

    Ok(CellValue::Float(total))
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
    async fn test_eval_sum_basic() {
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
        let result = eval_sum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 60.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_sum_empty() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_sum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_product_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(2));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(3));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(4));
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_product(ctx, "Sheet1", &args).await.unwrap();
        // 2 * 3 * 4 = 24
        match result {
            CellValue::Float(v) => assert!((v - 24.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_product_no_numbers() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::String("text".to_string()));
        engine.set_cell("Sheet1", 1, 0, CellValue::Empty);
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 1,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_product(ctx, "Sheet1", &args).await.unwrap();
        // PRODUCT returns 0 when no numeric values found
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_sumproduct_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Array 1: 1, 2, 3
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(1));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(2));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(3));
        // Array 2: 4, 5, 6
        engine.set_cell("Sheet1", 0, 1, CellValue::Int(4));
        engine.set_cell("Sheet1", 1, 1, CellValue::Int(5));
        engine.set_cell("Sheet1", 2, 1, CellValue::Int(6));
        let range1 = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let range2 = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 1,
            end_row: 2,
            end_col: 1,
        });
        let args = vec![range1, range2];
        let result = eval_sumproduct(ctx, "Sheet1", &args).await.unwrap();
        // 1*4 + 2*5 + 3*6 = 4 + 10 + 18 = 32
        match result {
            CellValue::Float(v) => assert!((v - 32.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_sumproduct_mismatched_dims() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Array 1: 1, 2, 3 (3 rows)
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(1));
        engine.set_cell("Sheet1", 1, 0, CellValue::Int(2));
        engine.set_cell("Sheet1", 2, 0, CellValue::Int(3));
        // Array 2: 4, 5 (2 rows)
        engine.set_cell("Sheet1", 0, 1, CellValue::Int(4));
        engine.set_cell("Sheet1", 1, 1, CellValue::Int(5));
        let range1 = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let range2 = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 1,
            end_row: 1,
            end_col: 1,
        });
        let args = vec![range1, range2];
        let result = eval_sumproduct(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#VALUE!")),
            _ => panic!("Expected Error, got {:?}", result),
        }
    }
}
