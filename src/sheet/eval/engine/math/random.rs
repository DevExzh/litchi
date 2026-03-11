use rand::Rng;

use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

pub(crate) async fn eval_rand(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("RAND expects 0 arguments".to_string()));
    }
    let mut rng = rand::rng();
    Ok(CellValue::Float(rng.random::<f64>()))
}

pub(crate) async fn eval_randbetween(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "RANDBETWEEN expects 2 arguments (bottom, top)".to_string(),
        ));
    }
    let bottom_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let top_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let bottom = match to_number(&bottom_val) {
        Some(n) => n.trunc() as i64,
        None => {
            return Ok(CellValue::Error(
                "RANDBETWEEN bottom must be numeric".to_string(),
            ));
        },
    };
    let top = match to_number(&top_val) {
        Some(n) => n.trunc() as i64,
        None => {
            return Ok(CellValue::Error(
                "RANDBETWEEN top must be numeric".to_string(),
            ));
        },
    };
    if bottom > top {
        return Ok(CellValue::Error(
            "RANDBETWEEN bottom must be less than or equal to top".to_string(),
        ));
    }
    let mut rng = rand::rng();
    let value = if bottom == top {
        bottom
    } else {
        rng.random_range(bottom..=top)
    };
    Ok(CellValue::Int(value))
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
    async fn test_eval_rand() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_rand(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => {
                // RAND should return a value between 0 and 1
                assert!(v >= 0.0 && v < 1.0);
            },
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_rand_with_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_rand(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 0 arguments")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_randbetween() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0), num_expr(10.0)];
        let result = eval_randbetween(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => {
                // RANDBETWEEN should return an integer in the range [1, 10]
                assert!(v >= 1 && v <= 10);
            },
            _ => panic!("Expected Int"),
        }
    }

    #[tokio::test]
    async fn test_eval_randbetween_same_value() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(5.0), num_expr(5.0)];
        let result = eval_randbetween(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // When bottom == top, should always return that value
            CellValue::Int(v) => assert_eq!(v, 5),
            _ => panic!("Expected Int"),
        }
    }

    #[tokio::test]
    async fn test_eval_randbetween_negative() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-10.0), num_expr(-1.0)];
        let result = eval_randbetween(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => {
                assert!(v >= -10 && v <= -1);
            },
            _ => panic!("Expected Int"),
        }
    }

    #[tokio::test]
    async fn test_eval_randbetween_mixed_range() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-5.0), num_expr(5.0)];
        let result = eval_randbetween(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => {
                assert!(v >= -5 && v <= 5);
            },
            _ => panic!("Expected Int"),
        }
    }

    #[tokio::test]
    async fn test_eval_randbetween_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_randbetween(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 arguments")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_randbetween_bottom_greater_than_top() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(10.0), num_expr(1.0)];
        let result = eval_randbetween(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("bottom must be less than or equal to top")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_randbetween_non_numeric() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            Expr::Literal(CellValue::String("abc".to_string())),
            num_expr(10.0),
        ];
        let result = eval_randbetween(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("bottom must be numeric")),
            _ => panic!("Expected Error"),
        }
    }
}
