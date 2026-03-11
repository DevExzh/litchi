use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, evaluate_expression, flatten_range_expr, to_bool, to_number};
use super::helpers::values_equal;

pub(crate) async fn eval_vlookup(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "VLOOKUP expects 3 or 4 arguments (lookup_value, table_array, col_index_num, [range_lookup])"
                .to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let table = flatten_range_expr(ctx, current_sheet, &args[1]).await?;

    let col_index_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
    let col_index = match to_number(&col_index_val) {
        Some(n) if n >= 1.0 => n as i64,
        _ => {
            return Ok(CellValue::Error(
                "VLOOKUP col_index_num must be a positive number".to_string(),
            ));
        },
    };

    let exact_match_only = if args.len() == 4 {
        let rl_val = evaluate_expression(ctx, current_sheet, &args[3]).await?;
        !to_bool(&rl_val)
    } else {
        true
    };

    if !exact_match_only {
        return Ok(CellValue::Error(
            "VLOOKUP currently only supports exact match (range_lookup = FALSE)".to_string(),
        ));
    }

    let rows = table.rows as i64;
    let cols = table.cols as i64;

    if col_index < 1 || col_index > cols {
        return Ok(CellValue::Error(
            "VLOOKUP col_index_num out of bounds for table_array".to_string(),
        ));
    }

    for r in 0..rows {
        let base = (r * cols) as usize;
        let key = &table.values[base];
        if values_equal(&lookup_val, key) {
            let idx = base + (col_index - 1) as usize;
            return Ok(table.values[idx].clone());
        }
    }

    Ok(CellValue::Error("VLOOKUP: value not found".to_string()))
}

pub(crate) async fn eval_hlookup(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "HLOOKUP expects 3 or 4 arguments (lookup_value, table_array, row_index_num, [range_lookup])"
                .to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let table = flatten_range_expr(ctx, current_sheet, &args[1]).await?;

    let row_index_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
    let row_index = match to_number(&row_index_val) {
        Some(n) if n >= 1.0 => n as i64,
        _ => {
            return Ok(CellValue::Error(
                "HLOOKUP row_index_num must be a positive number".to_string(),
            ));
        },
    };

    let exact_match_only = if args.len() == 4 {
        let rl_val = evaluate_expression(ctx, current_sheet, &args[3]).await?;
        !to_bool(&rl_val)
    } else {
        true
    };

    if !exact_match_only {
        return Ok(CellValue::Error(
            "HLOOKUP currently only supports exact match (range_lookup = FALSE)".to_string(),
        ));
    }

    let rows = table.rows as i64;
    let cols = table.cols as i64;

    if row_index < 1 || row_index > rows {
        return Ok(CellValue::Error(
            "HLOOKUP row_index_num out of bounds for table_array".to_string(),
        ));
    }

    for c in 0..cols {
        let key = &table.values[c as usize];
        if values_equal(&lookup_val, key) {
            let idx = ((row_index - 1) * cols + c) as usize;
            return Ok(table.values[idx].clone());
        }
    }

    Ok(CellValue::Error("HLOOKUP: value not found".to_string()))
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
    async fn test_vlookup_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        // Create a table: | ID | Name  | Value |
        //                   | 1  | Alice | 100   |
        //                   | 2  | Bob   | 200   |
        //                   | 3  | Carol | 300   |
        let values = vec![
            CellValue::Int(1),
            CellValue::String("Alice".to_string()),
            CellValue::Int(100),
            CellValue::Int(2),
            CellValue::String("Bob".to_string()),
            CellValue::Int(200),
            CellValue::Int(3),
            CellValue::String("Carol".to_string()),
            CellValue::Int(300),
        ];
        engine.add_range("Sheet1", 1, 1, 3, 3, values);

        // VLOOKUP(2, Sheet1!A1:C3, 2) should return "Bob"
        let args = vec![
            num_expr(2.0),
            range_expr("Sheet1", 1, 1, 3, 3),
            num_expr(2.0),
        ];
        let result = eval_vlookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("Bob".to_string()));
    }

    #[tokio::test]
    async fn test_vlookup_third_column() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![
            CellValue::Int(1),
            CellValue::String("Alice".to_string()),
            CellValue::Int(100),
            CellValue::Int(2),
            CellValue::String("Bob".to_string()),
            CellValue::Int(200),
        ];
        engine.add_range("Sheet1", 1, 1, 2, 3, values);

        // VLOOKUP(2, Sheet1!A1:C2, 3) should return 200
        let args = vec![
            num_expr(2.0),
            range_expr("Sheet1", 1, 1, 2, 3),
            num_expr(3.0),
        ];
        let result = eval_vlookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(200));
    }

    #[tokio::test]
    async fn test_vlookup_not_found() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![
            CellValue::Int(1),
            CellValue::String("Alice".to_string()),
            CellValue::Int(2),
            CellValue::String("Bob".to_string()),
        ];
        engine.add_range("Sheet1", 1, 1, 2, 2, values);

        let args = vec![
            num_expr(999.0),
            range_expr("Sheet1", 1, 1, 2, 2),
            num_expr(2.0),
        ];
        let result = eval_vlookup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("not found")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_vlookup_col_index_out_of_bounds() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![CellValue::Int(1), CellValue::String("Alice".to_string())];
        engine.add_range("Sheet1", 1, 1, 1, 2, values);

        let args = vec![
            num_expr(1.0),
            range_expr("Sheet1", 1, 1, 1, 2),
            num_expr(5.0),
        ];
        let result = eval_vlookup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("out of bounds")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_vlookup_string_lookup() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![
            CellValue::String("apple".to_string()),
            CellValue::Int(100),
            CellValue::String("banana".to_string()),
            CellValue::Int(200),
        ];
        engine.add_range("Sheet1", 1, 1, 2, 2, values);

        // VLOOKUP("banana", Sheet1!A1:B2, 2) should return 200
        let args = vec![
            str_expr("banana"),
            range_expr("Sheet1", 1, 1, 2, 2),
            num_expr(2.0),
        ];
        let result = eval_vlookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(200));
    }

    #[tokio::test]
    async fn test_hlookup_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        // Create a horizontal table:
        // | Name  | Alice | Bob   | Carol |
        // | Value | 100   | 200   | 300   |
        let values = vec![
            CellValue::String("Name".to_string()),
            CellValue::String("Alice".to_string()),
            CellValue::String("Bob".to_string()),
            CellValue::String("Carol".to_string()),
            CellValue::String("Value".to_string()),
            CellValue::Int(100),
            CellValue::Int(200),
            CellValue::Int(300),
        ];
        engine.add_range("Sheet1", 1, 1, 2, 4, values);

        // HLOOKUP("Bob", Sheet1!A1:D2, 2) should return 200
        let args = vec![
            str_expr("Bob"),
            range_expr("Sheet1", 1, 1, 2, 4),
            num_expr(2.0),
        ];
        let result = eval_hlookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(200));
    }

    #[tokio::test]
    async fn test_hlookup_not_found() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![
            CellValue::String("A".to_string()),
            CellValue::String("B".to_string()),
            CellValue::Int(1),
            CellValue::Int(2),
        ];
        engine.add_range("Sheet1", 1, 1, 2, 2, values);

        let args = vec![
            str_expr("Z"),
            range_expr("Sheet1", 1, 1, 2, 2),
            num_expr(2.0),
        ];
        let result = eval_hlookup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("not found")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_hlookup_row_index_out_of_bounds() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();

        let values = vec![CellValue::String("A".to_string()), CellValue::Int(1)];
        engine.add_range("Sheet1", 1, 1, 1, 2, values);

        let args = vec![
            str_expr("A"),
            range_expr("Sheet1", 1, 1, 1, 2),
            num_expr(5.0),
        ];
        let result = eval_hlookup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("out of bounds")),
            _ => panic!("Expected Error"),
        }
    }
}
