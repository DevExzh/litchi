use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use chrono::Datelike;

use super::helpers::{coerce_date_value, weekday_number, weeknum_value};

pub(crate) async fn eval_weekday(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "WEEKDAY expects 1 or 2 arguments (serial_number, [return_type])".to_string(),
        ));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let date = match coerce_date_value(&value) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "WEEKDAY expects a valid date serial or text".to_string(),
            ));
        },
    };
    let return_type = if args.len() == 2 {
        let rt_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&rt_val) {
            Some(n) if matches!(n as i32, 1..=3) => n as i32,
            _ => {
                return Ok(CellValue::Error(
                    "WEEKDAY return_type must be 1, 2, or 3".to_string(),
                ));
            },
        }
    } else {
        1
    };
    let number = match weekday_number(date.weekday(), return_type) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "WEEKDAY return_type not supported".to_string(),
            ));
        },
    };
    Ok(CellValue::Int(number))
}

pub(crate) async fn eval_weeknum(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "WEEKNUM expects 1 or 2 arguments (serial_number, [return_type])".to_string(),
        ));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let date = match coerce_date_value(&value) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "WEEKNUM expects a valid date serial or text".to_string(),
            ));
        },
    };
    let return_type = if args.len() == 2 {
        let rt_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&rt_val) {
            Some(n) if matches!(n as i32, 1 | 2) => n as i32,
            _ => {
                return Ok(CellValue::Error(
                    "WEEKNUM return_type must be 1 or 2".to_string(),
                ));
            },
        }
    } else {
        1
    };
    match weeknum_value(date, return_type) {
        Some(week) => Ok(CellValue::Int(week)),
        None => Ok(CellValue::Error(
            "WEEKNUM return_type not supported".to_string(),
        )),
    }
}

pub(crate) async fn eval_isoweeknum(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error(
            "ISOWEEKNUM expects 1 argument".to_string(),
        ));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let date = match coerce_date_value(&value) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "ISOWEEKNUM expects a valid date serial or text".to_string(),
            ));
        },
    };
    Ok(CellValue::Int(date.iso_week().week() as i64))
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

    // WEEKDAY tests
    #[tokio::test]
    async fn test_eval_weekday_default() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 2024-03-15 is a Friday
        let args = vec![num_expr(45366.0)];
        let result = eval_weekday(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // Default return_type=1: Sunday=1, Monday=2, ..., Friday=6
            CellValue::Int(v) => assert_eq!(v, 6),
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_weekday_type1() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 2024-03-15 is a Friday
        let args = vec![num_expr(45366.0), num_expr(1.0)];
        let result = eval_weekday(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // return_type=1: Sunday=1, Monday=2, ..., Friday=6
            CellValue::Int(v) => assert_eq!(v, 6),
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_weekday_type2() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 2024-03-15 is a Friday
        let args = vec![num_expr(45366.0), num_expr(2.0)];
        let result = eval_weekday(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // return_type=2: Monday=1, Tuesday=2, ..., Friday=5
            CellValue::Int(v) => assert_eq!(v, 5),
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_weekday_type3() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 2024-03-15 is a Friday
        let args = vec![num_expr(45366.0), num_expr(3.0)];
        let result = eval_weekday(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // return_type=3: Monday=0, Tuesday=1, ..., Friday=4
            CellValue::Int(v) => assert_eq!(v, 4),
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_weekday_sunday() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 2024-03-17 is a Sunday
        let args = vec![num_expr(45368.0), num_expr(1.0)];
        let result = eval_weekday(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // return_type=1: Sunday=1
            CellValue::Int(v) => assert_eq!(v, 1),
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_weekday_invalid_type() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(45366.0), num_expr(4.0)];
        let result = eval_weekday(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("return_type must be 1, 2, or 3")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_weekday_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_weekday(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 or 2 arguments")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_weekday_invalid_date() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("invalid".to_string()))];
        let result = eval_weekday(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects a valid date")),
            _ => panic!("Expected Error"),
        }
    }

    // WEEKNUM tests
    #[tokio::test]
    async fn test_eval_weeknum_default() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 2024-03-15 (Friday) - week starting Sunday
        let args = vec![num_expr(45366.0)];
        let result = eval_weeknum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => {
                // Week number should be positive
                assert!(v > 0 && v <= 53);
            },
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_weeknum_type1() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 2024-03-15 with return_type=1 (week starts Sunday)
        let args = vec![num_expr(45366.0), num_expr(1.0)];
        let result = eval_weeknum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => {
                assert!(v > 0 && v <= 53);
            },
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_weeknum_type2() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 2024-03-15 with return_type=2 (week starts Monday)
        let args = vec![num_expr(45366.0), num_expr(2.0)];
        let result = eval_weeknum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => {
                assert!(v > 0 && v <= 53);
            },
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_weeknum_invalid_type() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(45366.0), num_expr(3.0)];
        let result = eval_weeknum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("return_type must be 1 or 2")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_weeknum_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_weeknum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 or 2 arguments")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_weeknum_invalid_date() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("invalid".to_string()))];
        let result = eval_weeknum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects a valid date")),
            _ => panic!("Expected Error"),
        }
    }

    // ISOWEEKNUM tests
    #[tokio::test]
    async fn test_eval_isoweeknum_basic() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 2024-03-15
        let args = vec![num_expr(45366.0)];
        let result = eval_isoweeknum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => {
                // ISO week number should be between 1 and 53
                assert!((1..=53).contains(&v));
            },
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_isoweeknum_january() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 2024-01-01 (Monday) - first day of year, ISO week 1
        let args = vec![num_expr(45292.0)];
        let result = eval_isoweeknum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => {
                assert_eq!(v, 1);
            },
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_isoweeknum_year_boundary() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // 2023-12-31 - last day of 2023, could be week 52 or 1 of next year
        // Serial for 2023-12-31: 45291
        let args = vec![num_expr(45291.0)];
        let result = eval_isoweeknum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => {
                // Could be 52 (most years) or 1 (if it's a Monday of week 1 of next year)
                assert!(v == 52 || v == 1);
            },
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_isoweeknum_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_isoweeknum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 argument")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_isoweeknum_invalid_date() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("invalid".to_string()))];
        let result = eval_isoweeknum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects a valid date")),
            _ => panic!("Expected Error"),
        }
    }
}
