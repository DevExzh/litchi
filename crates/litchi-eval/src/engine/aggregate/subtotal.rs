use crate::parser::Expr;
use litchi_core::sheet::{CellValue, Result};

use super::super::statistical::{eval_stdev_p, eval_stdev_s, eval_var_p, eval_var_s};
use super::super::{EvalCtx, evaluate_expression, to_number};
use super::{eval_average, eval_count, eval_counta, eval_max, eval_min, eval_product, eval_sum};

/// SUBTOTAL(function_num, ref1, [ref2], ...) — aggregates the given
/// references with the function selected by `function_num`:
/// 1=AVERAGE, 2=COUNT, 3=COUNTA, 4=MAX, 5=MIN, 6=PRODUCT, 7=STDEV,
/// 8=STDEVP, 9=SUM, 10=VAR, 11=VARP.
///
/// Limitations:
/// - Codes 101..111 are the hidden-row-aware variants in Excel. The
///   `EngineCtx`/`WorkbookTrait` surface does not expose row visibility, so
///   hidden rows cannot be excluded; 101..111 currently aggregate the same
///   values as 1..11.
/// - Excel also ignores cells containing nested SUBTOTAL formulas inside the
///   referenced ranges. Cells arrive here as plain values, so nested
///   subtotals cannot be detected and are counted like any other value.
pub(crate) async fn eval_subtotal(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 {
        return Ok(CellValue::Error(
            "SUBTOTAL expects at least 2 arguments (function_num, ref1, ...)".to_string(),
        ));
    }

    let num_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = num_val {
        return Ok(num_val);
    }
    let function_num = match to_number(&num_val) {
        Some(n) => n.trunc() as i64,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let base = match function_num {
        1..=11 => function_num,
        101..=111 => function_num - 100,
        _ => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let sub_args = &args[1..];
    match base {
        1 => eval_average(ctx, current_sheet, sub_args).await,
        2 => eval_count(ctx, current_sheet, sub_args).await,
        3 => eval_counta(ctx, current_sheet, sub_args).await,
        4 => eval_max(ctx, current_sheet, sub_args).await,
        5 => eval_min(ctx, current_sheet, sub_args).await,
        6 => eval_product(ctx, current_sheet, sub_args).await,
        7 => eval_stdev_s(ctx, current_sheet, sub_args).await,
        8 => eval_stdev_p(ctx, current_sheet, sub_args).await,
        9 => eval_sum(ctx, current_sheet, sub_args).await,
        10 => eval_var_s(ctx, current_sheet, sub_args).await,
        11 => eval_var_p(ctx, current_sheet, sub_args).await,
        _ => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::engine::test_helpers::TestEngine;
    use crate::parser::ast::RangeRef;

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

    fn sample_engine() -> TestEngine {
        let engine = TestEngine::new();
        engine.add_range(
            "Sheet1",
            1,
            1,
            4,
            1,
            vec![
                CellValue::Int(1),
                CellValue::Int(2),
                CellValue::Int(3),
                CellValue::Int(4),
            ],
        );
        engine
    }

    #[tokio::test]
    async fn test_subtotal_sum() {
        let engine = sample_engine();
        let ctx = engine.ctx();
        let args = vec![num_expr(9.0), range_expr("Sheet1", 1, 1, 4, 1)];
        let result = eval_subtotal(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 10.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_subtotal_average() {
        let engine = sample_engine();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0), range_expr("Sheet1", 1, 1, 4, 1)];
        let result = eval_subtotal(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 2.5).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_subtotal_count_and_counta() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.add_range(
            "Sheet1",
            1,
            1,
            3,
            1,
            vec![
                CellValue::Int(5),
                CellValue::String("text".to_string()),
                CellValue::Empty,
            ],
        );
        let range = range_expr("Sheet1", 1, 1, 3, 1);
        // COUNT (2) counts only numbers.
        let result = eval_subtotal(ctx, "Sheet1", &[num_expr(2.0), range.clone()])
            .await
            .unwrap();
        assert_eq!(result, CellValue::Int(1));
        // COUNTA (3) counts all non-empty cells.
        let result = eval_subtotal(ctx, "Sheet1", &[num_expr(3.0), range])
            .await
            .unwrap();
        assert_eq!(result, CellValue::Int(2));
    }

    #[tokio::test]
    async fn test_subtotal_max_min_product() {
        let engine = sample_engine();
        let ctx = engine.ctx();
        let range = range_expr("Sheet1", 1, 1, 4, 1);
        match eval_subtotal(ctx, "Sheet1", &[num_expr(4.0), range.clone()])
            .await
            .unwrap()
        {
            CellValue::Float(v) => assert!((v - 4.0).abs() < 1e-9),
            other => panic!("Expected 4, got {:?}", other),
        }
        match eval_subtotal(ctx, "Sheet1", &[num_expr(5.0), range.clone()])
            .await
            .unwrap()
        {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            other => panic!("Expected 1, got {:?}", other),
        }
        match eval_subtotal(ctx, "Sheet1", &[num_expr(6.0), range])
            .await
            .unwrap()
        {
            CellValue::Float(v) => assert!((v - 24.0).abs() < 1e-9),
            other => panic!("Expected 24, got {:?}", other),
        }
    }

    #[tokio::test]
    async fn test_subtotal_stdev_and_variants() {
        let engine = sample_engine();
        let ctx = engine.ctx();
        let range = range_expr("Sheet1", 1, 1, 4, 1);
        // STDEV (7): sample stdev of 1..4
        match eval_subtotal(ctx, "Sheet1", &[num_expr(7.0), range.clone()])
            .await
            .unwrap()
        {
            CellValue::Float(v) => assert!((v - 1.290994).abs() < 1e-5),
            other => panic!("Expected Float, got {:?}", other),
        }
        // STDEVP (8)
        match eval_subtotal(ctx, "Sheet1", &[num_expr(8.0), range.clone()])
            .await
            .unwrap()
        {
            CellValue::Float(v) => assert!((v - 1.118034).abs() < 1e-5),
            other => panic!("Expected Float, got {:?}", other),
        }
        // VAR (10)
        match eval_subtotal(ctx, "Sheet1", &[num_expr(10.0), range.clone()])
            .await
            .unwrap()
        {
            CellValue::Float(v) => assert!((v - 1.666667).abs() < 1e-5),
            other => panic!("Expected Float, got {:?}", other),
        }
        // VARP (11)
        match eval_subtotal(ctx, "Sheet1", &[num_expr(11.0), range])
            .await
            .unwrap()
        {
            CellValue::Float(v) => assert!((v - 1.25).abs() < 1e-5),
            other => panic!("Expected Float, got {:?}", other),
        }
    }

    #[tokio::test]
    async fn test_subtotal_hidden_aware_codes() {
        // 101..111 are accepted and aggregate like 1..11 because row
        // visibility is not exposed by the engine trait (see doc comment).
        let engine = sample_engine();
        let ctx = engine.ctx();
        let args = vec![num_expr(109.0), range_expr("Sheet1", 1, 1, 4, 1)];
        let result = eval_subtotal(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 10.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_subtotal_invalid_function_num() {
        let engine = sample_engine();
        let ctx = engine.ctx();
        for bad in [0.0, 12.0, 100.0, 112.0, -1.0] {
            let args = vec![num_expr(bad), range_expr("Sheet1", 1, 1, 4, 1)];
            let result = eval_subtotal(ctx, "Sheet1", &args).await.unwrap();
            match result {
                CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
                _ => panic!("Expected #VALUE! for {}, got {:?}", bad, result),
            }
        }
    }

    #[tokio::test]
    async fn test_subtotal_non_numeric_function_num() {
        let engine = sample_engine();
        let ctx = engine.ctx();
        let args = vec![
            Expr::Literal(CellValue::String("sum".to_string())),
            range_expr("Sheet1", 1, 1, 4, 1),
        ];
        let result = eval_subtotal(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            _ => panic!("Expected #VALUE!, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_subtotal_error_function_num_propagates() {
        let engine = sample_engine();
        let ctx = engine.ctx();
        let args = vec![
            Expr::Literal(CellValue::Error("#REF!".to_string())),
            range_expr("Sheet1", 1, 1, 4, 1),
        ];
        let result = eval_subtotal(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#REF!"),
            _ => panic!("Expected #REF!, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_subtotal_too_few_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(9.0)];
        let result = eval_subtotal(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects at least 2")),
            _ => panic!("Expected Error"),
        }
    }
}
