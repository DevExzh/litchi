use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_text};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

pub(crate) async fn eval_unichar(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("UNICHAR expects 1 argument".to_string()));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let Some(n) = crate::sheet::eval::engine::to_number(&val) {
        let code = n.trunc() as u32;
        if code == 0 {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }
        match char::from_u32(code) {
            Some(c) => Ok(CellValue::String(c.to_string())),
            None => Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        Ok(CellValue::Error("#VALUE!".to_string()))
    }
}

pub(crate) async fn eval_unicode(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("UNICODE expects 1 argument".to_string()));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&val);
    if let Some(c) = text.chars().next() {
        Ok(CellValue::Int(c as u32 as i64))
    } else {
        Ok(CellValue::Error("#VALUE!".to_string()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::engine::test_helpers::TestEngine;
    use crate::sheet::eval::parser::Expr;

    #[tokio::test]
    async fn test_eval_unichar_ascii() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Int(65))]; // 'A'
        let result = eval_unichar(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("A".to_string()));
    }

    #[tokio::test]
    async fn test_eval_unichar_space() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Int(32))]; // space
        let result = eval_unichar(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String(" ".to_string()));
    }

    #[tokio::test]
    async fn test_eval_unichar_zero() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Int(0))];
        let result = eval_unichar(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Error("#VALUE!".to_string()));
    }

    #[tokio::test]
    async fn test_eval_unichar_invalid() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("abc".to_string()))];
        let result = eval_unichar(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Error("#VALUE!".to_string()));
    }

    #[tokio::test]
    async fn test_eval_unichar_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_unichar(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 argument")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_unicode_ascii() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("A".to_string()))];
        let result = eval_unicode(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(65));
    }

    #[tokio::test]
    async fn test_eval_unicode_space() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String(" ".to_string()))];
        let result = eval_unicode(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(32));
    }

    #[tokio::test]
    async fn test_eval_unicode_empty() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("".to_string()))];
        let result = eval_unicode(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Error("#VALUE!".to_string()));
    }

    #[tokio::test]
    async fn test_eval_unicode_first_char() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Should return code for first character only
        let args = vec![Expr::Literal(CellValue::String("ABC".to_string()))];
        let result = eval_unicode(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(65));
    }

    #[tokio::test]
    async fn test_eval_unicode_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_unicode(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 argument")),
            _ => panic!("Expected Error"),
        }
    }
}
