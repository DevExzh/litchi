use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, flatten_range_expr, to_number};

pub(crate) async fn eval_index(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "INDEX expects 2 or 3 arguments (array, row_num, [column_num])".to_string(),
        ));
    }

    let array = flatten_range_expr(ctx, current_sheet, &args[0]).await?;

    let row_val = super::super::evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let row_num = match to_number(&row_val) {
        Some(n) if n >= 1.0 => n as i64,
        _ => {
            return Ok(CellValue::Error(
                "INDEX row_num must be a positive number".to_string(),
            ));
        },
    };

    let col_num = if args.len() == 3 {
        let col_val = super::super::evaluate_expression(ctx, current_sheet, &args[2]).await?;
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::engine::test_helpers::TestEngine;
    use crate::sheet::eval::parser::Expr;
    use crate::sheet::eval::parser::ast::RangeRef;

    fn num_expr(n: f64) -> Expr {
        if n == n.floor() {
            Expr::Literal(CellValue::Int(n as i64))
        } else {
            Expr::Literal(CellValue::Float(n))
        }
    }

    fn range_expr(sheet: &str, start_row: u32, start_col: u32, end_row: u32, end_col: u32) -> Expr {
        Expr::Range(RangeRef {
            sheet: sheet.to_string(),
            start_row,
            start_col,
            end_row,
            end_col,
        })
    }

    #[tokio::test]
    async fn test_index_with_range() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        // Set up a 2x3 range with values
        let values = vec![
            CellValue::Int(1),
            CellValue::Int(2),
            CellValue::Int(3),
            CellValue::Int(4),
            CellValue::Int(5),
            CellValue::Int(6),
        ];
        engine.add_range("Sheet1", 1, 1, 2, 3, values);

        // Test INDEX(Sheet1!A1:C2, 2, 3) should return 6
        let args = vec![
            range_expr("Sheet1", 1, 1, 2, 3),
            num_expr(2.0),
            num_expr(3.0),
        ];
        let result = eval_index(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(6));
    }

    #[tokio::test]
    async fn test_index_single_row() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![CellValue::Int(10), CellValue::Int(20), CellValue::Int(30)];
        engine.add_range("Sheet1", 1, 1, 1, 3, values);

        // INDEX(Sheet1!A1:C1, 1, 2) should return 20
        let args = vec![
            range_expr("Sheet1", 1, 1, 1, 3),
            num_expr(1.0),
            num_expr(2.0),
        ];
        let result = eval_index(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(20));
    }

    #[tokio::test]
    async fn test_index_default_column() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![
            CellValue::Int(1),
            CellValue::Int(2),
            CellValue::Int(3),
            CellValue::Int(4),
        ];
        engine.add_range("Sheet1", 1, 1, 2, 2, values);

        // INDEX(Sheet1!A1:B2, 2) should default to column 1, returning 3
        let args = vec![range_expr("Sheet1", 1, 1, 2, 2), num_expr(2.0)];
        let result = eval_index(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(3));
    }

    #[tokio::test]
    async fn test_index_wrong_number_of_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let args = vec![num_expr(1.0)];
        let result = eval_index(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 or 3")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_index_zero_row() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![CellValue::Int(1)];
        engine.add_range("Sheet1", 1, 1, 1, 1, values);

        let args = vec![range_expr("Sheet1", 1, 1, 1, 1), num_expr(0.0)];
        let result = eval_index(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("must be a positive number")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_index_out_of_bounds() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![CellValue::Int(1), CellValue::Int(2)];
        engine.add_range("Sheet1", 1, 1, 1, 2, values);

        let args = vec![range_expr("Sheet1", 1, 1, 1, 2), num_expr(5.0)];
        let result = eval_index(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("out of bounds")),
            _ => panic!("Expected Error"),
        }
    }
}
