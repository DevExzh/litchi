use crate::sheet::eval::engine::EvalCtx;
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use chrono::Utc;

use super::helpers::{date_to_excel_serial_1900, datetime_to_excel_serial_1900};

pub(crate) async fn eval_today(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("TODAY expects 0 arguments".to_string()));
    }

    let now = Utc::now().date_naive();
    let serial = date_to_excel_serial_1900(now);
    Ok(CellValue::DateTime(serial))
}

pub(crate) async fn eval_now(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("NOW expects 0 arguments".to_string()));
    }

    let now = Utc::now().naive_utc();
    let serial = datetime_to_excel_serial_1900(now);
    Ok(CellValue::DateTime(serial))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::parser::Expr;

    #[tokio::test]
    async fn test_eval_today() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_today(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::DateTime(v) => {
                // TODAY should return a date serial (integer part only)
                assert!(v > 0.0);
                // Should be a whole number (no time component)
                assert_eq!(v, v.floor());
            },
            _ => panic!("Expected DateTime"),
        }
    }

    #[tokio::test]
    async fn test_eval_today_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Int(1))];
        let result = eval_today(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 0 arguments")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_now() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_now(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::DateTime(v) => {
                // NOW should return a datetime serial with fractional part
                assert!(v > 0.0);
            },
            _ => panic!("Expected DateTime"),
        }
    }

    #[tokio::test]
    async fn test_eval_now_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Int(1))];
        let result = eval_now(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 0 arguments")),
            _ => panic!("Expected Error"),
        }
    }
}
