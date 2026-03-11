use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, evaluate_expression, flatten_range_expr, to_number};
use super::helpers::{find_exact_match_index, is_1d};

pub(crate) async fn eval_xlookup(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 6 {
        return Ok(CellValue::Error(
            "XLOOKUP expects 3 to 6 arguments (lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])"
                .to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let lookup_array = flatten_range_expr(ctx, current_sheet, &args[1]).await?;
    let return_array = flatten_range_expr(ctx, current_sheet, &args[2]).await?;

    if !is_1d(&lookup_array) || !is_1d(&return_array) {
        return Ok(CellValue::Error(
            "XLOOKUP lookup_array and return_array must be one-dimensional ranges".to_string(),
        ));
    }

    if lookup_array.values.len() != return_array.values.len() {
        return Ok(CellValue::Error(
            "XLOOKUP lookup_array and return_array must have the same length".to_string(),
        ));
    }

    let if_not_found = if args.len() >= 4 {
        Some(evaluate_expression(ctx, current_sheet, &args[3]).await?)
    } else {
        None
    };

    if args.len() >= 5 {
        let mm_val = evaluate_expression(ctx, current_sheet, &args[4]).await?;
        match to_number(&mm_val) {
            Some(0.0) => {},
            _ => {
                return Ok(CellValue::Error(
                    "XLOOKUP currently only supports match_mode = 0 (exact match)".to_string(),
                ));
            },
        }
    }

    if args.len() == 6 {
        return Ok(CellValue::Error(
            "XLOOKUP search_mode is not supported in this evaluator".to_string(),
        ));
    }

    if let Some(idx) = find_exact_match_index(&lookup_val, &lookup_array.values) {
        Ok(return_array.values[idx].clone())
    } else if let Some(not_found) = if_not_found {
        Ok(not_found)
    } else {
        Ok(CellValue::Error("XLOOKUP: value not found".to_string()))
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
    async fn test_xlookup_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        // Create lookup array [1, 2, 3] and return array ["A", "B", "C"]
        let lookup_values = vec![CellValue::Int(1), CellValue::Int(2), CellValue::Int(3)];
        let return_values = vec![
            CellValue::String("A".to_string()),
            CellValue::String("B".to_string()),
            CellValue::String("C".to_string()),
        ];
        engine.add_range("Sheet1", 1, 1, 1, 3, lookup_values);
        engine.add_range("Sheet1", 2, 1, 1, 3, return_values);

        // XLOOKUP(2, Sheet1!A1:C1, Sheet1!A2:C2) should return "B"
        let args = vec![
            num_expr(2.0),
            range_expr("Sheet1", 1, 1, 1, 3),
            range_expr("Sheet1", 2, 1, 2, 3),
        ];
        let result = eval_xlookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("B".to_string()));
    }

    #[tokio::test]
    async fn test_xlookup_not_found() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let lookup_values = vec![CellValue::Int(1), CellValue::Int(2)];
        let return_values = vec![
            CellValue::String("A".to_string()),
            CellValue::String("B".to_string()),
        ];
        engine.add_range("Sheet1", 1, 1, 1, 2, lookup_values);
        engine.add_range("Sheet1", 2, 1, 1, 2, return_values);

        let args = vec![
            num_expr(99.0),
            range_expr("Sheet1", 1, 1, 1, 2),
            range_expr("Sheet1", 2, 1, 2, 2),
        ];
        let result = eval_xlookup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("not found")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_xlookup_if_not_found() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let lookup_values = vec![CellValue::Int(1), CellValue::Int(2)];
        let return_values = vec![
            CellValue::String("A".to_string()),
            CellValue::String("B".to_string()),
        ];
        engine.add_range("Sheet1", 1, 1, 1, 2, lookup_values);
        engine.add_range("Sheet1", 2, 1, 1, 2, return_values);

        // XLOOKUP(99, lookup, return, "Not Found")
        let args = vec![
            num_expr(99.0),
            range_expr("Sheet1", 1, 1, 1, 2),
            range_expr("Sheet1", 2, 1, 2, 2),
            str_expr("Not Found"),
        ];
        let result = eval_xlookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("Not Found".to_string()));
    }

    #[tokio::test]
    async fn test_xlookup_string_lookup() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let lookup_values = vec![
            CellValue::String("apple".to_string()),
            CellValue::String("banana".to_string()),
            CellValue::String("cherry".to_string()),
        ];
        let return_values = vec![
            CellValue::Int(100),
            CellValue::Int(200),
            CellValue::Int(300),
        ];
        engine.add_range("Sheet1", 1, 1, 1, 3, lookup_values);
        engine.add_range("Sheet1", 2, 1, 1, 3, return_values);

        // XLOOKUP("banana", lookup, return) should return 200
        let args = vec![
            str_expr("banana"),
            range_expr("Sheet1", 1, 1, 1, 3),
            range_expr("Sheet1", 2, 1, 2, 3),
        ];
        let result = eval_xlookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(200));
    }

    #[tokio::test]
    async fn test_xlookup_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let args = vec![num_expr(1.0), num_expr(2.0)];
        let result = eval_xlookup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 3 to 6")),
            _ => panic!("Expected Error"),
        }
    }
}
