use crate::sheet::eval::engine::EvalCtx;
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::helpers::{
    add_months, date_to_excel_serial_1900, last_day_of_month, number_arg, serial_to_excel_date_1900,
};

pub(crate) async fn eval_edate(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "EDATE expects 2 arguments (start_date, months)".to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "EDATE start_date is not numeric".to_string(),
            ));
        },
    };
    let months = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("EDATE months is not numeric".to_string()));
        },
    };

    let date = match serial_to_excel_date_1900(start) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "EDATE start_date is not a valid date".to_string(),
            ));
        },
    };

    let shifted = match add_months(date, months.trunc() as i32) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error("EDATE result is out of range".to_string()));
        },
    };

    let serial = date_to_excel_serial_1900(shifted);
    Ok(CellValue::DateTime(serial))
}

pub(crate) async fn eval_eomonth(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "EOMONTH expects 2 arguments (start_date, months)".to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "EOMONTH start_date is not numeric".to_string(),
            ));
        },
    };
    let months = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "EOMONTH months is not numeric".to_string(),
            ));
        },
    };

    let date = match serial_to_excel_date_1900(start) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "EOMONTH start_date is not a valid date".to_string(),
            ));
        },
    };

    let shifted = match add_months(date, months.trunc() as i32) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "EOMONTH result is out of range".to_string(),
            ));
        },
    };

    let last_day = match last_day_of_month(shifted) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "EOMONTH result is out of range".to_string(),
            ));
        },
    };

    let serial = date_to_excel_serial_1900(last_day);
    Ok(CellValue::DateTime(serial))
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

    #[tokio::test]
    async fn test_eval_edate_add_months() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Start from 2024-03-15 (serial ~45366)
        let args = vec![num_expr(45366.0), num_expr(2.0)];
        let result = eval_edate(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::DateTime(v) => {
                // Should be around 2024-05-15
                assert!(v > 45366.0);
            },
            _ => panic!("Expected DateTime"),
        }
    }

    #[tokio::test]
    async fn test_eval_edate_subtract_months() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(45366.0), num_expr(-2.0)];
        let result = eval_edate(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::DateTime(v) => {
                // Should be around 2024-01-15
                assert!(v < 45366.0);
            },
            _ => panic!("Expected DateTime"),
        }
    }

    #[tokio::test]
    async fn test_eval_edate_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(45366.0)];
        let result = eval_edate(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 arguments")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_edate_invalid_date() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Use a very large negative serial that would cause date underflow
        let args = vec![num_expr(-1_000_000_000.0), num_expr(2.0)];
        let result = eval_edate(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("not a valid date")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_eomonth_add_months() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Start from 2024-03-15 (serial ~45366), add 2 months -> 2024-05-31
        let args = vec![num_expr(45366.0), num_expr(2.0)];
        let result = eval_eomonth(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::DateTime(v) => {
                // Should be 2024-05-31
                assert!(v > 45366.0);
            },
            _ => panic!("Expected DateTime"),
        }
    }

    #[tokio::test]
    async fn test_eval_eomonth_zero_months() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Start from 2024-03-15, get end of current month (2024-03-31)
        let args = vec![num_expr(45366.0), num_expr(0.0)];
        let result = eval_eomonth(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::DateTime(v) => {
                // Should be 2024-03-31
                assert!(v > 45366.0);
            },
            _ => panic!("Expected DateTime"),
        }
    }

    #[tokio::test]
    async fn test_eval_eomonth_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(45366.0)];
        let result = eval_eomonth(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 arguments")),
            _ => panic!("Expected Error"),
        }
    }
}
