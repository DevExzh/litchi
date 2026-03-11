use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, for_each_value_in_expr, is_blank};

pub(crate) async fn eval_count(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut count = 0u64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if matches!(
                v,
                CellValue::Int(_) | CellValue::Float(_) | CellValue::DateTime(_)
            ) {
                count += 1;
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Int(count as i64))
}

pub(crate) async fn eval_countblank(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut count = 0u64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if is_blank(v) {
                count += 1;
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Int(count as i64))
}

pub(crate) async fn eval_counta(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut count = 0u64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if !matches!(v, CellValue::Empty) {
                count += 1;
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Int(count as i64))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::engine::test_helpers::TestEngine;
    use crate::sheet::eval::parser::Expr;

    #[tokio::test]
    async fn test_eval_count_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(10));
        engine.set_cell("Sheet1", 1, 0, CellValue::Float(20.5));
        engine.set_cell("Sheet1", 2, 0, CellValue::String("text".to_string()));
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_count(ctx, "Sheet1", &args).await.unwrap();
        // COUNT only counts numeric values (Int, Float, DateTime)
        assert_eq!(result, CellValue::Int(2));
    }

    #[tokio::test]
    async fn test_eval_count_empty() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_count(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(0));
    }

    #[tokio::test]
    async fn test_eval_counta_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(10));
        engine.set_cell("Sheet1", 1, 0, CellValue::String("text".to_string()));
        engine.set_cell("Sheet1", 2, 0, CellValue::Empty);
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_counta(ctx, "Sheet1", &args).await.unwrap();
        // COUNTA counts all non-empty cells
        assert_eq!(result, CellValue::Int(2));
    }

    #[tokio::test]
    async fn test_eval_countblank_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::Int(10));
        engine.set_cell("Sheet1", 1, 0, CellValue::Empty);
        engine.set_cell("Sheet1", 2, 0, CellValue::Empty);
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        });
        let args = vec![range];
        let result = eval_countblank(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(2));
    }
}
