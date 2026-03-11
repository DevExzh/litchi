use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, evaluate_expression, flatten_range_expr, to_number};
use super::helpers::{find_exact_match_index, is_1d};

pub(crate) async fn eval_match(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "MATCH expects 2 or 3 arguments (lookup_value, lookup_array, [match_type])".to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let lookup_array = flatten_range_expr(ctx, current_sheet, &args[1]).await?;

    if !is_1d(&lookup_array) {
        return Ok(CellValue::Error(
            "MATCH lookup_array must be a one-dimensional range".to_string(),
        ));
    }

    let match_type = if args.len() == 3 {
        let mt_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match to_number(&mt_val) {
            Some(0.0) => 0,
            _ => {
                return Ok(CellValue::Error(
                    "MATCH currently only supports match_type = 0 (exact match)".to_string(),
                ));
            },
        }
    } else {
        0
    };

    if match_type != 0 {
        return Ok(CellValue::Error(
            "MATCH currently only supports exact match (match_type = 0)".to_string(),
        ));
    }

    if let Some(idx) = find_exact_match_index(&lookup_val, &lookup_array.values) {
        Ok(CellValue::Int((idx + 1) as i64))
    } else {
        Ok(CellValue::Error("MATCH: value not found".to_string()))
    }
}

pub(crate) async fn eval_xmatch(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 4 {
        return Ok(CellValue::Error(
            "XMATCH expects 2 to 4 arguments (lookup_value, lookup_array, [match_mode], [search_mode])"
                .to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let lookup_array = flatten_range_expr(ctx, current_sheet, &args[1]).await?;

    if !is_1d(&lookup_array) {
        return Ok(CellValue::Error(
            "XMATCH lookup_array must be a one-dimensional range".to_string(),
        ));
    }

    if args.len() >= 3 {
        let mm_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match to_number(&mm_val) {
            Some(0.0) => {},
            _ => {
                return Ok(CellValue::Error(
                    "XMATCH currently only supports match_mode = 0 (exact match)".to_string(),
                ));
            },
        }
    }

    if args.len() == 4 {
        return Ok(CellValue::Error(
            "XMATCH search_mode is not supported in this evaluator".to_string(),
        ));
    }

    if let Some(idx) = find_exact_match_index(&lookup_val, &lookup_array.values) {
        Ok(CellValue::Int((idx + 1) as i64))
    } else {
        Ok(CellValue::Error("XMATCH: value not found".to_string()))
    }
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

    fn str_expr(s: &str) -> Expr {
        Expr::Literal(CellValue::String(s.to_string()))
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
    async fn test_eval_match_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        // Set up a lookup array [10, 20, 30, 40, 50]
        let values = vec![
            CellValue::Int(10),
            CellValue::Int(20),
            CellValue::Int(30),
            CellValue::Int(40),
            CellValue::Int(50),
        ];
        engine.add_range("Sheet1", 1, 1, 1, 5, values);

        // MATCH(30, Sheet1!A1:E1) should return 3
        let args = vec![num_expr(30.0), range_expr("Sheet1", 1, 1, 1, 5)];
        let result = eval_match(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(3));
    }

    #[tokio::test]
    async fn test_eval_match_string() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![
            CellValue::String("apple".to_string()),
            CellValue::String("banana".to_string()),
            CellValue::String("cherry".to_string()),
        ];
        engine.add_range("Sheet1", 1, 1, 1, 3, values);

        // MATCH("banana", Sheet1!A1:C1) should return 2
        let args = vec![str_expr("banana"), range_expr("Sheet1", 1, 1, 1, 3)];
        let result = eval_match(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(2));
    }

    #[tokio::test]
    async fn test_eval_match_not_found() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![CellValue::Int(1), CellValue::Int(2), CellValue::Int(3)];
        engine.add_range("Sheet1", 1, 1, 1, 3, values);

        let args = vec![num_expr(99.0), range_expr("Sheet1", 1, 1, 1, 3)];
        let result = eval_match(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("not found")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_match_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let args = vec![num_expr(1.0)];
        let result = eval_match(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 or 3")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_xmatch_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![
            CellValue::Int(100),
            CellValue::Int(200),
            CellValue::Int(300),
        ];
        engine.add_range("Sheet1", 1, 1, 1, 3, values);

        // XMATCH(200, Sheet1!A1:C1) should return 2
        let args = vec![num_expr(200.0), range_expr("Sheet1", 1, 1, 1, 3)];
        let result = eval_xmatch(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(2));
    }

    #[tokio::test]
    async fn test_eval_xmatch_not_found() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![CellValue::Int(1), CellValue::Int(2)];
        engine.add_range("Sheet1", 1, 1, 1, 2, values);

        let args = vec![num_expr(999.0), range_expr("Sheet1", 1, 1, 1, 2)];
        let result = eval_xmatch(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("not found")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_xmatch_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let args = vec![num_expr(1.0)];
        let result = eval_xmatch(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 to 4")),
            _ => panic!("Expected Error"),
        }
    }
}
