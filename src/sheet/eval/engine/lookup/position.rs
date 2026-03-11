use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, flatten_range_expr};
use super::helpers::{ReferenceLookup, first_cell_from_expr};

pub(crate) async fn eval_column(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() > 1 {
        return Ok(CellValue::Error(
            "COLUMN expects at most 1 argument".to_string(),
        ));
    }

    let col = if args.is_empty() {
        match ctx.current_position() {
            Some((_, _, col)) => col,
            None => {
                return Ok(CellValue::Error(
                    "COLUMN cannot determine the current position".to_string(),
                ));
            },
        }
    } else {
        match first_cell_from_expr(ctx, current_sheet, &args[0]).await? {
            ReferenceLookup::Point((_row, col)) => col,
            ReferenceLookup::NameError(msg) => {
                return Ok(CellValue::Error(msg));
            },
            ReferenceLookup::NotReference => {
                return Ok(CellValue::Error(
                    "COLUMN expects a cell or range reference".to_string(),
                ));
            },
        }
    };

    Ok(CellValue::Int(col as i64))
}

pub(crate) async fn eval_row(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() > 1 {
        return Ok(CellValue::Error(
            "ROW expects at most 1 argument".to_string(),
        ));
    }

    let row = if args.is_empty() {
        match ctx.current_position() {
            Some((_, row, _)) => row,
            None => {
                return Ok(CellValue::Error(
                    "ROW cannot determine the current position".to_string(),
                ));
            },
        }
    } else {
        match first_cell_from_expr(ctx, current_sheet, &args[0]).await? {
            ReferenceLookup::Point((row, _col)) => row,
            ReferenceLookup::NameError(msg) => {
                return Ok(CellValue::Error(msg));
            },
            ReferenceLookup::NotReference => {
                return Ok(CellValue::Error(
                    "ROW expects a cell or range reference".to_string(),
                ));
            },
        }
    };

    Ok(CellValue::Int(row as i64))
}

pub(crate) async fn eval_rows(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error(
            "ROWS expects exactly 1 argument (array)".to_string(),
        ));
    }

    let range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    Ok(CellValue::Int(range.rows as i64))
}

pub(crate) async fn eval_columns(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error(
            "COLUMNS expects exactly 1 argument (array)".to_string(),
        ));
    }

    let range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    Ok(CellValue::Int(range.cols as i64))
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
    async fn test_eval_column_with_reference() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        // Set current position to row 5, col 3 (C5)
        engine.set_current_position("Sheet1", 5, 3);

        // COLUMN() without args should return current column (3)
        let args: Vec<Expr> = vec![];
        let result = eval_column(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(3));
    }

    #[tokio::test]
    async fn test_eval_row_with_reference() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        // Set current position to row 5, col 3 (C5)
        engine.set_current_position("Sheet1", 5, 3);

        // ROW() without args should return current row (5)
        let args: Vec<Expr> = vec![];
        let result = eval_row(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(5));
    }

    #[tokio::test]
    async fn test_eval_rows() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![
            CellValue::Int(1),
            CellValue::Int(2),
            CellValue::Int(3),
            CellValue::Int(4),
            CellValue::Int(5),
            CellValue::Int(6),
        ];
        engine.add_range("Sheet1", 1, 1, 3, 2, values);

        // ROWS(Sheet1!A1:B3) should return 3
        let args = vec![range_expr("Sheet1", 1, 1, 3, 2)];
        let result = eval_rows(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(3));
    }

    #[tokio::test]
    async fn test_eval_columns() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![
            CellValue::Int(1),
            CellValue::Int(2),
            CellValue::Int(3),
            CellValue::Int(4),
            CellValue::Int(5),
            CellValue::Int(6),
        ];
        engine.add_range("Sheet1", 1, 1, 2, 3, values);

        // COLUMNS(Sheet1!A1:C2) should return 3
        let args = vec![range_expr("Sheet1", 1, 1, 2, 3)];
        let result = eval_columns(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(3));
    }

    #[tokio::test]
    async fn test_eval_rows_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let args = vec![num_expr(1.0), num_expr(2.0)];
        let result = eval_rows(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects exactly 1")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_columns_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let args: Vec<Expr> = vec![];
        let result = eval_columns(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects exactly 1")),
            _ => panic!("Expected Error"),
        }
    }
}
