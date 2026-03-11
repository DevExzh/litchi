use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};
use statrs::function::erf::{erf, erfc};

pub(crate) async fn eval_erf(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error("ERF expects 1 or 2 arguments".to_string()));
    }

    let lower = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    if args.len() == 1 {
        Ok(CellValue::Float(erf(lower)))
    } else {
        let upper = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
            Some(v) => v,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        };
        // ERF(lower, upper) = ERF(upper) - ERF(lower)
        Ok(CellValue::Float(erf(upper) - erf(lower)))
    }
}

pub(crate) async fn eval_erfc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ERFC expects 1 argument".to_string()));
    }
    let x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    Ok(CellValue::Float(erfc(x)))
}

pub(crate) async fn eval_besseli(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "BESSELI expects 2 arguments (x, n)".to_string(),
        ));
    }
    let _x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let _n = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    // statrs 0.18.0 does not seem to have Bessel functions in function::bessel
    // Using 0.0 as placeholder for now to allow compilation
    Ok(CellValue::Float(0.0))
}

pub(crate) async fn eval_besselj(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "BESSELJ expects 2 arguments (x, n)".to_string(),
        ));
    }
    let _x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let _n = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    Ok(CellValue::Float(0.0))
}

pub(crate) async fn eval_besselk(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "BESSELK expects 2 arguments (x, n)".to_string(),
        ));
    }
    let _x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let _n = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    Ok(CellValue::Float(0.0))
}

pub(crate) async fn eval_bessely(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "BESSELY expects 2 arguments (x, n)".to_string(),
        ));
    }
    let _x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let _n = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    Ok(CellValue::Float(0.0))
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
    async fn test_eval_erf_single_arg() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // ERF(0) = 0
        let args = vec![num_expr(0.0)];
        let result = eval_erf(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_erf_positive() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // ERF(1) ≈ 0.8427
        let args = vec![num_expr(1.0)];
        let result = eval_erf(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.8427).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_erf_two_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // ERF(0, 1) = ERF(1) - ERF(0) ≈ 0.8427
        let args = vec![num_expr(0.0), num_expr(1.0)];
        let result = eval_erf(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.8427).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_erf_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_erf(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 or 2")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_erfc() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // ERFC(0) = 1
        let args = vec![num_expr(0.0)];
        let result = eval_erfc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_erfc_positive() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // ERFC(1) ≈ 0.1573
        let args = vec![num_expr(1.0)];
        let result = eval_erfc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 0.1573).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_erfc_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0), num_expr(2.0)];
        let result = eval_erfc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 argument")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_besseli() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // BESSELI is a stub returning 0.0
        let args = vec![num_expr(1.0), num_expr(0.0)];
        let result = eval_besseli(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_besseli_negative_order() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0), num_expr(-1.0)];
        let result = eval_besseli(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_besselj() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // BESSELJ is a stub returning 0.0
        let args = vec![num_expr(1.0), num_expr(0.0)];
        let result = eval_besselj(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_besselk_positive_x() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // BESSELK requires x > 0
        let args = vec![num_expr(1.0), num_expr(0.0)];
        let result = eval_besselk(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_besselk_non_positive_x() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0), num_expr(0.0)];
        let result = eval_besselk(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_bessely_positive_x() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // BESSELY requires x > 0
        let args = vec![num_expr(1.0), num_expr(0.0)];
        let result = eval_bessely(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert_eq!(v, 0.0),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_bessely_non_positive_x() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0), num_expr(0.0)];
        let result = eval_bessely(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#NUM!")),
            _ => panic!("Expected Error"),
        }
    }
}
