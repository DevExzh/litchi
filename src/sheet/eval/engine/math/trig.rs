use std::f64::consts::FRAC_PI_2;

use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::helpers::EPS;

pub(crate) async fn eval_sin(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "SIN", f64::sin).await
}

pub(crate) async fn eval_cos(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "COS", f64::cos).await
}

pub(crate) async fn eval_tan(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "TAN", f64::tan).await
}

pub(crate) async fn eval_cot(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("COT expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let angle = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("COT on non-numeric value".to_string())),
    };
    let tan = angle.tan();
    if tan.abs() < EPS {
        return Ok(CellValue::Error(
            "COT is undefined for angles where TAN(angle) = 0".to_string(),
        ));
    }
    Ok(CellValue::Float(1.0 / tan))
}

pub(crate) async fn eval_acot(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ACOT", |x| FRAC_PI_2 - x.atan()).await
}

pub(crate) async fn eval_csc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("CSC expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let angle = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("CSC on non-numeric value".to_string())),
    };
    let sin = angle.sin();
    if sin.abs() < EPS {
        return Ok(CellValue::Error(
            "CSC is undefined for angles where SIN(angle) = 0".to_string(),
        ));
    }
    Ok(CellValue::Float(1.0 / sin))
}

pub(crate) async fn eval_sec(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("SEC expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let angle = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("SEC on non-numeric value".to_string())),
    };
    let cos = angle.cos();
    if cos.abs() < EPS {
        return Ok(CellValue::Error(
            "SEC is undefined for angles where COS(angle) = 0".to_string(),
        ));
    }
    Ok(CellValue::Float(1.0 / cos))
}

pub(crate) async fn eval_sinh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "SINH", f64::sinh).await
}

pub(crate) async fn eval_cosh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "COSH", f64::cosh).await
}

pub(crate) async fn eval_tanh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "TANH", f64::tanh).await
}

pub(crate) async fn eval_csch(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("CSCH expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let value = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("CSCH on non-numeric value".to_string())),
    };
    let sinh = value.sinh();
    if sinh.abs() < EPS {
        return Ok(CellValue::Error(
            "CSCH is undefined for values where SINH(value) = 0".to_string(),
        ));
    }
    Ok(CellValue::Float(1.0 / sinh))
}

pub(crate) async fn eval_sech(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("SECH expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let value = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("SECH on non-numeric value".to_string())),
    };
    let cosh = value.cosh();
    if cosh.abs() < EPS {
        return Ok(CellValue::Error(
            "SECH is undefined for values where COSH(value) = 0".to_string(),
        ));
    }
    Ok(CellValue::Float(1.0 / cosh))
}

pub(crate) async fn eval_coth(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("COTH expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let value = match to_number(&v) {
        Some(n) => n,
        None => return Ok(CellValue::Error("COTH on non-numeric value".to_string())),
    };
    let sinh = value.sinh();
    if sinh.abs() < EPS {
        return Ok(CellValue::Error(
            "COTH is undefined for values where SINH(value) = 0".to_string(),
        ));
    }
    Ok(CellValue::Float(value.cosh() / sinh))
}

pub(crate) async fn eval_asin(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ASIN", f64::asin).await
}

pub(crate) async fn eval_acos(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ACOS", f64::acos).await
}

pub(crate) async fn eval_atan(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ATAN", f64::atan).await
}

pub(crate) async fn eval_atan2(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("ATAN2 expects 2 arguments".to_string()));
    }
    let y = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let yn = match to_number(&y) {
        Some(n) => n,
        None => return Ok(CellValue::Error("ATAN2 y is not numeric".to_string())),
    };
    let xn = match to_number(&x) {
        Some(n) => n,
        None => return Ok(CellValue::Error("ATAN2 x is not numeric".to_string())),
    };
    Ok(CellValue::Float(yn.atan2(xn)))
}

pub(crate) async fn eval_asinh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ASINH", f64::asinh).await
}

pub(crate) async fn eval_acosh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ACOSH", f64::acosh).await
}

pub(crate) async fn eval_atanh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "ATANH", f64::atanh).await
}

pub(crate) async fn eval_acoth(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ACOTH expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = match to_number(&v) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "ACOTH requires a numeric argument".to_string(),
            ));
        },
    };
    if x.abs() <= 1.0 {
        return Ok(CellValue::Error(
            "ACOTH input must satisfy |x| > 1".to_string(),
        ));
    }
    let value = 0.5 * ((x + 1.0) / (x - 1.0)).ln();
    Ok(CellValue::Float(value))
}

pub(crate) async fn eval_radians(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "RADIANS", f64::to_radians).await
}

pub(crate) async fn eval_degrees(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_trig(ctx, current_sheet, args, "DEGREES", f64::to_degrees).await
}

async fn eval_unary_trig<F>(
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
        return Ok(CellValue::Error(format!("{name} expects 1 argument")));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let n = match to_number(&v) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(format!(
                "{name} on non-numeric value: {:?}",
                v
            )));
        },
    };
    Ok(CellValue::Float(f(n)))
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
    async fn test_eval_sin() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_sin(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sin_pi_over_2() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(std::f64::consts::FRAC_PI_2)];
        let result = eval_sin(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_cos() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_cos(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_cos_pi() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(std::f64::consts::PI)];
        let result = eval_cos(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - (-1.0)).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_tan() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_tan(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_tan_pi_over_4() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(std::f64::consts::FRAC_PI_4)];
        let result = eval_tan(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_cot() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(std::f64::consts::FRAC_PI_4)];
        let result = eval_cot(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_cot_undefined() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_cot(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("undefined")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_csc() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(std::f64::consts::FRAC_PI_2)];
        let result = eval_csc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_csc_undefined() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_csc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("undefined")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sec() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_sec(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sec_undefined() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(std::f64::consts::FRAC_PI_2)];
        let result = eval_sec(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("undefined")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_asin() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_asin(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_asin_one() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_asin(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::FRAC_PI_2).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_acos() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_acos(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_acos_zero() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_acos(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::FRAC_PI_2).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_atan() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_atan(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_atan_one() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_atan(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::FRAC_PI_4).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_atan2() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0), num_expr(1.0)];
        let result = eval_atan2(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::FRAC_PI_4).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_atan2_zero_x() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0), num_expr(0.0)];
        let result = eval_atan2(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::FRAC_PI_2).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sinh() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_sinh(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_cosh() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_cosh(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_tanh() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_tanh(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_csch() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_csch(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0 / 1.0_f64.sinh()).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_csch_undefined() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_csch(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("undefined")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sech() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_sech(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_coth() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_coth(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => {
                let expected = 1.0_f64.cosh() / 1.0_f64.sinh();
                assert!((v - expected).abs() < 1e-9)
            },
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_coth_undefined() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_coth(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("undefined")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_acot() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_acot(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::FRAC_PI_4).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_asinh() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_asinh(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_acosh() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_acosh(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_atanh() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_atanh(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_acoth() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(2.0)];
        let result = eval_acoth(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => {
                let expected = 0.5 * ((2.0_f64 + 1.0) / (2.0 - 1.0)).ln();
                assert!((v - expected).abs() < 1e-9)
            },
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_acoth_invalid() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.5)];
        let result = eval_acoth(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("|x| > 1")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_radians() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(180.0)];
        let result = eval_radians(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::PI).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_degrees() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(std::f64::consts::PI)];
        let result = eval_degrees(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 180.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sin_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_sin(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 argument")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sin_non_numeric() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("abc".to_string()))];
        let result = eval_sin(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("non-numeric")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_atan2_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_atan2(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 arguments")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_atan2_non_numeric_y() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            Expr::Literal(CellValue::String("abc".to_string())),
            num_expr(1.0),
        ];
        let result = eval_atan2(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("y is not numeric")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_atan2_non_numeric_x() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(1.0),
            Expr::Literal(CellValue::String("abc".to_string())),
        ];
        let result = eval_atan2(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("x is not numeric")),
            _ => panic!("Expected Error result"),
        }
    }
}
