use crate::sheet::eval::engine::{EvalCtx, evaluate_expression};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use chrono::Datelike;

use super::helpers::{SECONDS_PER_DAY, coerce_date_value, coerce_time_fraction};

pub(crate) async fn eval_year(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("YEAR expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match coerce_date_value(&value) {
        Some(date) => Ok(CellValue::Int(date.year() as i64)),
        None => Ok(CellValue::Error(
            "YEAR expects a valid date serial or text".to_string(),
        )),
    }
}

pub(crate) async fn eval_month(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("MONTH expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match coerce_date_value(&value) {
        Some(date) => Ok(CellValue::Int(date.month() as i64)),
        None => Ok(CellValue::Error(
            "MONTH expects a valid date serial or text".to_string(),
        )),
    }
}

pub(crate) async fn eval_day(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("DAY expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match coerce_date_value(&value) {
        Some(date) => Ok(CellValue::Int(date.day() as i64)),
        None => Ok(CellValue::Error(
            "DAY expects a valid date serial or text".to_string(),
        )),
    }
}

pub(crate) async fn eval_hour(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("HOUR expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match coerce_time_fraction(&value) {
        Some(frac) => {
            let seconds = (frac * SECONDS_PER_DAY).rem_euclid(SECONDS_PER_DAY);
            Ok(CellValue::Int((seconds / 3600.0).floor() as i64))
        },
        None => Ok(CellValue::Error(
            "HOUR expects a valid time serial or text".to_string(),
        )),
    }
}

pub(crate) async fn eval_minute(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("MINUTE expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match coerce_time_fraction(&value) {
        Some(frac) => {
            let seconds = (frac * SECONDS_PER_DAY).rem_euclid(SECONDS_PER_DAY);
            let minutes = (seconds / 60.0).floor() as i64 % 60;
            Ok(CellValue::Int(minutes))
        },
        None => Ok(CellValue::Error(
            "MINUTE expects a valid time serial or text".to_string(),
        )),
    }
}

pub(crate) async fn eval_second(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("SECOND expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match coerce_time_fraction(&value) {
        Some(frac) => {
            let seconds = (frac * SECONDS_PER_DAY).rem_euclid(SECONDS_PER_DAY);
            Ok(CellValue::Int((seconds % 60.0).round() as i64))
        },
        None => Ok(CellValue::Error(
            "SECOND expects a valid time serial or text".to_string(),
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::parser::Expr;

    fn num_expr(n: f64) -> Expr {
        if n == n.floor() {
            Expr::Literal(CellValue::Int(n as i64))
        } else {
            Expr::Literal(CellValue::Float(n))
        }
    }

    #[tokio::test]
    async fn test_eval_year_from_serial() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Excel serial for 2024-03-15 is approximately 45366
        let args = vec![num_expr(45366.0)];
        let result = eval_year(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 2024),
            _ => panic!("Expected Int(2024)"),
        }
    }

    #[tokio::test]
    async fn test_eval_month_from_serial() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Excel serial for 2024-03-15
        let args = vec![num_expr(45366.0)];
        let result = eval_month(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 3),
            _ => panic!("Expected Int(3)"),
        }
    }

    #[tokio::test]
    async fn test_eval_day_from_serial() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Excel serial for 2024-03-15
        let args = vec![num_expr(45366.0)];
        let result = eval_day(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 15),
            _ => panic!("Expected Int(15)"),
        }
    }

    #[tokio::test]
    async fn test_eval_hour() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // 0.5 is 12:00:00 (noon)
        let args = vec![num_expr(0.5)];
        let result = eval_hour(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 12),
            _ => panic!("Expected Int(12)"),
        }
    }

    #[tokio::test]
    async fn test_eval_hour_with_date() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // 45366.75 is 2024-03-15 18:00:00
        let args = vec![num_expr(45366.75)];
        let result = eval_hour(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 18),
            _ => panic!("Expected Int(18)"),
        }
    }

    #[tokio::test]
    async fn test_eval_minute() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // 12:30:00 = (12 * 3600 + 30 * 60) / 86400 = 45000 / 86400 = 0.520833333333...
        let args = vec![num_expr(45000.0 / 86400.0)];
        let result = eval_minute(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 30),
            _ => panic!("Expected Int(30)"),
        }
    }

    #[tokio::test]
    async fn test_eval_second() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // 0.000011574 is approximately 1 second (1/86400)
        let args = vec![num_expr(0.000011574)];
        let result = eval_second(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 1),
            _ => panic!("Expected Int(1)"),
        }
    }

    #[tokio::test]
    async fn test_eval_year_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_year(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 argument")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_hour_invalid() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("not a time".to_string()))];
        let result = eval_hour(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects a valid time")),
            _ => panic!("Expected Error"),
        }
    }
}
