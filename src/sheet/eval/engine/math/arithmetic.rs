use crate::sheet::{CellValue, Result};

use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;

pub(crate) async fn eval_int(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("INT expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let n = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("INT on non-numeric value".to_string())),
    };
    Ok(CellValue::Int(n.floor() as i64))
}

pub(crate) async fn eval_abs(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ABS expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match v {
        CellValue::Int(i) => Ok(CellValue::Int(i.abs())),
        CellValue::Float(f) => Ok(CellValue::Float(f.abs())),
        other => Ok(CellValue::Error(format!(
            "ABS on non-numeric value: {:?}",
            other
        ))),
    }
}

pub(crate) async fn eval_power(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("POWER expects 2 arguments".to_string()));
    }
    let base = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let exp = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let b = match to_number(&base) {
        Some(n) => n,
        None => return Ok(CellValue::Error("POWER base is not numeric".to_string())),
    };
    let e = match to_number(&exp) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "POWER exponent is not numeric".to_string(),
            ));
        },
    };
    Ok(CellValue::Float(b.powf(e)))
}

pub(crate) async fn eval_ln(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "LN", |x| x.ln()).await
}

pub(crate) async fn eval_log10(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "LOG10", |x| x.log10()).await
}

pub(crate) async fn eval_log(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 && args.len() != 2 {
        return Ok(CellValue::Error(
            "LOG expects 1 or 2 arguments (number, [base])".to_string(),
        ));
    }
    let number = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let n = match to_number(&number) {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("LOG number must be positive".to_string())),
        None => return Ok(CellValue::Error("LOG number is not numeric".to_string())),
    };
    let base = if args.len() == 2 {
        let base_expr = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&base_expr) {
            Some(b) if b > 0.0 && b != 1.0 => b,
            Some(_) => {
                return Ok(CellValue::Error(
                    "LOG base must be positive and not equal to 1".to_string(),
                ));
            },
            None => return Ok(CellValue::Error("LOG base is not numeric".to_string())),
        }
    } else {
        10.0
    };
    Ok(CellValue::Float(n.log(base)))
}

pub(crate) async fn eval_exp(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "EXP", |x| x.exp()).await
}

pub(crate) async fn eval_sqrt(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "SQRT", |x| x.sqrt()).await
}

pub(crate) async fn eval_sqrtpi(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_math(ctx, current_sheet, args, "SQRTPI", |x| {
        (x * std::f64::consts::PI).sqrt()
    })
    .await
}

pub(crate) async fn eval_delta(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "DELTA expects 1 or 2 arguments (number, [comparison])".to_string(),
        ));
    }

    let first = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = match to_number(&first) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("DELTA number must be numeric".to_string()));
        },
    };

    let y = if args.len() == 2 {
        let second = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&second) {
            Some(n) => n,
            None => {
                return Ok(CellValue::Error(
                    "DELTA comparison value must be numeric".to_string(),
                ));
            },
        }
    } else {
        0.0
    };

    Ok(CellValue::Int(if x == y { 1 } else { 0 }))
}

pub(crate) async fn eval_gestep(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "GESTEP expects 1 or 2 arguments (number, [step])".to_string(),
        ));
    }

    let number_expr = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let number = match to_number(&number_expr) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "GESTEP number must be numeric".to_string(),
            ));
        },
    };

    let step = if args.len() == 2 {
        let step_expr = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&step_expr) {
            Some(n) => n,
            None => {
                return Ok(CellValue::Error("GESTEP step must be numeric".to_string()));
            },
        }
    } else {
        0.0
    };

    Ok(CellValue::Int(if number >= step { 1 } else { 0 }))
}

async fn eval_unary_math<F>(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    name: &str,
    f: F,
) -> Result<CellValue>
where
    F: Fn(f64) -> f64,
{
    if args.len() != 1 {
        return Ok(CellValue::Error(format!("{} expects 1 argument", name)));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let n = match to_number(&v) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(format!(
                "{} on non-numeric value: {:?}",
                name, v
            )));
        },
    };
    Ok(CellValue::Float(f(n)))
}

#[cfg(test)]
mod tests {
    use super::*;

    // Helper function to create a simple numeric expression
    fn num_expr(n: f64) -> Expr {
        if n == n.floor() {
            Expr::Literal(CellValue::Int(n as i64))
        } else {
            Expr::Literal(CellValue::Float(n))
        }
    }

    #[tokio::test]
    async fn test_eval_abs_int() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-42.0)];
        let result = eval_abs(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 42),
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_abs_float() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-std::f64::consts::PI)];
        let result = eval_abs(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::PI).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_power() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(2.0), num_expr(3.0)];
        let result = eval_power(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 8.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sqrt() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(16.0)];
        let result = eval_sqrt(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 4.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_ln() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(std::f64::consts::E)];
        let result = eval_ln(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_log10() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1000.0)];
        let result = eval_log10(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_exp() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_exp(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::E).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_int() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.7)];
        let result = eval_int(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 3),
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_int_negative() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-3.7)];
        let result = eval_int(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, -4),
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_delta_equal() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(5.0)];
        let result = eval_delta(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 0), // 5 != 0
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_delta_with_comparison() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(5.0), num_expr(5.0)];
        let result = eval_delta(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 1), // 5 == 5
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_gestep() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(5.0), num_expr(3.0)];
        let result = eval_gestep(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 1), // 5 >= 3
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_gestep_default() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(5.0)];
        let result = eval_gestep(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 1), // 5 >= 0
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sqrtpi() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(2.0)];
        let result = eval_sqrtpi(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - (2.0 * std::f64::consts::PI).sqrt()).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }
}
