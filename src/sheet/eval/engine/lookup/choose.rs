use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, evaluate_expression};

pub(crate) async fn eval_choose(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 {
        return Ok(CellValue::Error(
            "CHOOSE expects at least 2 arguments (index_num, value1, ...)".to_string(),
        ));
    }

    let index_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let index = match super::super::to_number(&index_val) {
        Some(n) if n >= 1.0 => n.trunc() as usize,
        _ => {
            return Ok(CellValue::Error(
                "CHOOSE index_num must be a positive number".to_string(),
            ));
        },
    };

    let choices = args.len() - 1;
    if index == 0 || index > choices {
        return Ok(CellValue::Error(
            "CHOOSE index_num out of range".to_string(),
        ));
    }

    let choice_expr = &args[index];
    evaluate_expression(ctx, current_sheet, choice_expr).await
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

    fn str_expr(s: &str) -> Expr {
        Expr::Literal(CellValue::String(s.to_string()))
    }

    #[tokio::test]
    async fn test_eval_choose_basic() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // CHOOSE(2, "a", "b", "c") should return "b"
        let args = vec![num_expr(2.0), str_expr("a"), str_expr("b"), str_expr("c")];
        let result = eval_choose(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "b"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_choose_first() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0), num_expr(10.0), num_expr(20.0)];
        let result = eval_choose(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(10.0) | CellValue::Int(10) => {},
            _ => panic!("Expected 10"),
        }
    }

    #[tokio::test]
    async fn test_eval_choose_last() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.0), str_expr("x"), str_expr("y"), str_expr("z")];
        let result = eval_choose(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "z"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_choose_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_choose(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects at least 2")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_choose_index_zero() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0), str_expr("a"), str_expr("b")];
        let result = eval_choose(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("must be a positive number")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_choose_index_out_of_range() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(5.0), str_expr("a"), str_expr("b")];
        let result = eval_choose(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("out of range")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_choose_non_numeric_index() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("invalid"), str_expr("a"), str_expr("b")];
        let result = eval_choose(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("must be a positive number")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_choose_decimal_index() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Decimal index should be truncated
        let args = vec![num_expr(2.9), str_expr("a"), str_expr("b"), str_expr("c")];
        let result = eval_choose(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "b"),
            _ => panic!("Expected String"),
        }
    }
}
