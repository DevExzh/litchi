use crate::sheet::{CellValue, Result};

use super::super::parser::Expr;
use super::{EvalCtx, evaluate_expression, for_each_value_in_expr, to_bool};

pub(crate) async fn eval_if(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error("IF expects 2 or 3 arguments".to_string()));
    }

    // Condition
    let cond_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let cond = to_bool(&cond_val);

    if cond {
        evaluate_expression(ctx, current_sheet, &args[1]).await
    } else if args.len() == 3 {
        evaluate_expression(ctx, current_sheet, &args[2]).await
    } else {
        Ok(CellValue::Bool(false))
    }
}

pub(crate) async fn eval_ifs(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || !args.len().is_multiple_of(2) {
        return Ok(CellValue::Error(
            "IFS expects an even number of arguments (condition/result pairs)".to_string(),
        ));
    }
    for pair in args.chunks(2) {
        let condition = evaluate_expression(ctx, current_sheet, &pair[0]).await?;
        if let CellValue::Error(_) = condition {
            return Ok(condition);
        }
        if to_bool(&condition) {
            return evaluate_expression(ctx, current_sheet, &pair[1]).await;
        }
    }
    Ok(CellValue::Error("#N/A".to_string()))
}

pub(crate) async fn eval_and(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Bool(true));
    }
    for arg in args {
        let mut all_true = true;
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if !to_bool(v) {
                all_true = false;
            }
            Ok(())
        })
        .await?;
        if !all_true {
            return Ok(CellValue::Bool(false));
        }
    }
    Ok(CellValue::Bool(true))
}

pub(crate) async fn eval_or(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Bool(false));
    }
    let mut any_true = false;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if to_bool(v) {
                any_true = true;
            }
            Ok(())
        })
        .await?;
        if any_true {
            return Ok(CellValue::Bool(true));
        }
    }
    Ok(CellValue::Bool(false))
}

pub(crate) async fn eval_xor(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "XOR expects at least 1 argument".to_string(),
        ));
    }
    let mut parity = false;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if to_bool(v) {
                parity = !parity;
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Bool(parity))
}

pub(crate) async fn eval_not(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("NOT expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    Ok(CellValue::Bool(!to_bool(&v)))
}

pub(crate) async fn eval_true(_: EvalCtx<'_>, _: &str, args: &[Expr]) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("TRUE expects no arguments".to_string()));
    }
    Ok(CellValue::Bool(true))
}

pub(crate) async fn eval_false(_: EvalCtx<'_>, _: &str, args: &[Expr]) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("FALSE expects no arguments".to_string()));
    }
    Ok(CellValue::Bool(false))
}

pub(crate) async fn eval_switch(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 {
        return Ok(CellValue::Error(
            "SWITCH expects at least 3 arguments (expression, value1, result1, ...)".to_string(),
        ));
    }

    let target = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = target {
        return Ok(target);
    }

    let mut i = 1;
    while i + 1 < args.len() {
        let val = evaluate_expression(ctx, current_sheet, &args[i]).await?;
        if val == target {
            return evaluate_expression(ctx, current_sheet, &args[i + 1]).await;
        }
        i += 2;
    }

    // Default value if exists (odd number of arguments total)
    if args.len().is_multiple_of(2) {
        evaluate_expression(ctx, current_sheet, &args[args.len() - 1]).await
    } else {
        Ok(CellValue::Error("#N/A".to_string()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::parser::Expr;

    fn bool_expr(b: bool) -> Expr {
        Expr::Literal(CellValue::Bool(b))
    }

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
    async fn test_eval_if_true() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true), num_expr(1.0), num_expr(0.0)];
        let result = eval_if(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 1),
            _ => panic!("Expected Int(1)"),
        }
    }

    #[tokio::test]
    async fn test_eval_if_false() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(false), num_expr(1.0), num_expr(0.0)];
        let result = eval_if(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 0),
            _ => panic!("Expected Int(0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_if_no_else() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(false), num_expr(1.0)];
        let result = eval_if(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(!v),
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_if_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true)];
        let result = eval_if(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 or 3")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_ifs_first_match() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            bool_expr(true),
            num_expr(1.0),
            bool_expr(false),
            num_expr(2.0),
        ];
        let result = eval_ifs(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 1),
            _ => panic!("Expected Int(1)"),
        }
    }

    #[tokio::test]
    async fn test_eval_ifs_second_match() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            bool_expr(false),
            num_expr(1.0),
            bool_expr(true),
            num_expr(2.0),
        ];
        let result = eval_ifs(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 2),
            _ => panic!("Expected Int(2)"),
        }
    }

    #[tokio::test]
    async fn test_eval_ifs_no_match() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            bool_expr(false),
            num_expr(1.0),
            bool_expr(false),
            num_expr(2.0),
        ];
        let result = eval_ifs(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#N/A"),
            _ => panic!("Expected #N/A error"),
        }
    }

    #[tokio::test]
    async fn test_eval_ifs_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true)];
        let result = eval_ifs(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("even number")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_and_empty() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_and(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_and_all_true() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true), bool_expr(true), bool_expr(true)];
        let result = eval_and(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_and_one_false() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true), bool_expr(false), bool_expr(true)];
        let result = eval_and(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(!v),
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_and_numeric_truthy() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0), num_expr(2.0)];
        let result = eval_and(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_and_numeric_falsy() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0), num_expr(0.0)];
        let result = eval_and(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(!v),
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_or_empty() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_or(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(!v),
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_or_one_true() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(false), bool_expr(true), bool_expr(false)];
        let result = eval_or(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_or_all_false() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(false), bool_expr(false), bool_expr(false)];
        let result = eval_or(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(!v),
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_or_numeric() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0), num_expr(1.0)];
        let result = eval_or(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_xor_empty() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_xor(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects at least 1")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_xor_one_true() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true)];
        let result = eval_xor(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_xor_two_true() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true), bool_expr(true)];
        let result = eval_xor(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(!v),
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_xor_three_true() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true), bool_expr(true), bool_expr(true)];
        let result = eval_xor(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_xor_mixed() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true), bool_expr(false), bool_expr(true)];
        let result = eval_xor(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(!v),
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_not_true() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true)];
        let result = eval_not(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(!v),
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_not_false() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(false)];
        let result = eval_not(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_not_zero() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_not(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_not_nonzero() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(5.0)];
        let result = eval_not(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(!v),
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_not_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_not(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 argument")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_true() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_true(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_true_with_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_true(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects no arguments")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_false() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_false(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(!v),
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_false_with_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_false(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects no arguments")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_switch_match_first() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(1.0),
            num_expr(1.0),
            str_expr("one"),
            num_expr(2.0),
            str_expr("two"),
        ];
        let result = eval_switch(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "one"),
            _ => panic!("Expected String('one')"),
        }
    }

    #[tokio::test]
    async fn test_eval_switch_match_second() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(2.0),
            num_expr(1.0),
            str_expr("one"),
            num_expr(2.0),
            str_expr("two"),
        ];
        let result = eval_switch(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "two"),
            _ => panic!("Expected String('two')"),
        }
    }

    #[tokio::test]
    async fn test_eval_switch_no_match() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(3.0),
            num_expr(1.0),
            str_expr("one"),
            num_expr(2.0),
            str_expr("two"),
        ];
        let result = eval_switch(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#N/A"),
            _ => panic!("Expected #N/A error"),
        }
    }

    #[tokio::test]
    async fn test_eval_switch_default() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(3.0),
            num_expr(1.0),
            str_expr("one"),
            num_expr(2.0),
            str_expr("two"),
            str_expr("default"),
        ];
        let result = eval_switch(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "default"),
            _ => panic!("Expected String('default')"),
        }
    }

    #[tokio::test]
    async fn test_eval_switch_too_few_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0), num_expr(1.0)];
        let result = eval_switch(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("at least 3")),
            _ => panic!("Expected Error"),
        }
    }
}
