use crate::sheet::{CellValue, Result};

use super::helpers::{EPS, is_even, number_result, round_away_from_zero};
use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;

pub(crate) async fn eval_round(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("ROUND expects 2 arguments".to_string()));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let digits = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => return Ok(CellValue::Error("ROUND number is not numeric".to_string())),
    };
    let d = match to_number(&digits) {
        Some(n) => n as i32,
        None => return Ok(CellValue::Error("ROUND digits is not numeric".to_string())),
    };
    let factor = 10f64.powi(d.abs());
    let result = if d >= 0 {
        (x * factor).round() / factor
    } else {
        (x / factor).round() * factor
    };
    Ok(CellValue::Float(result))
}

pub(crate) async fn eval_rounddown(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "ROUNDDOWN expects 2 arguments".to_string(),
        ));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let digits = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "ROUNDDOWN number is not numeric".to_string(),
            ));
        },
    };
    let d = match to_number(&digits) {
        Some(n) => n as i32,
        None => {
            return Ok(CellValue::Error(
                "ROUNDDOWN digits is not numeric".to_string(),
            ));
        },
    };
    let factor = 10f64.powi(d.abs());
    let result = if d >= 0 {
        (x * factor).trunc() / factor
    } else {
        (x / factor).trunc() * factor
    };
    Ok(CellValue::Float(result))
}

pub(crate) async fn eval_roundup(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("ROUNDUP expects 2 arguments".to_string()));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let digits = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "ROUNDUP number is not numeric".to_string(),
            ));
        },
    };
    let d = match to_number(&digits) {
        Some(n) => n as i32,
        None => {
            return Ok(CellValue::Error(
                "ROUNDUP digits is not numeric".to_string(),
            ));
        },
    };
    let factor = 10f64.powi(d.abs());
    let scaled = if d >= 0 { x * factor } else { x / factor };
    let rounded = if scaled >= 0.0 {
        scaled.ceil()
    } else {
        scaled.floor()
    };
    let result = if d >= 0 {
        rounded / factor
    } else {
        rounded * factor
    };
    Ok(CellValue::Float(result))
}

pub(crate) async fn eval_floor(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "FLOOR expects 1 or 2 arguments (number, [significance])".to_string(),
        ));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => return Ok(CellValue::Error("FLOOR number is not numeric".to_string())),
    };
    let sig = if args.len() == 2 {
        let s_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&s_val) {
            Some(n) if n != 0.0 => n,
            _ => {
                return Ok(CellValue::Error(
                    "FLOOR significance must be non-zero numeric".to_string(),
                ));
            },
        }
    } else {
        1.0
    };
    let result = (x / sig).floor() * sig;
    Ok(CellValue::Float(result))
}

pub(crate) async fn eval_ceiling(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "CEILING expects 1 or 2 arguments (number, [significance])".to_string(),
        ));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "CEILING number is not numeric".to_string(),
            ));
        },
    };
    let sig = if args.len() == 2 {
        let s_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&s_val) {
            Some(n) if n != 0.0 => n,
            _ => {
                return Ok(CellValue::Error(
                    "CEILING significance must be non-zero numeric".to_string(),
                ));
            },
        }
    } else {
        1.0
    };
    let result = (x / sig).ceil() * sig;
    Ok(CellValue::Float(result))
}

pub(crate) async fn eval_floor_math(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 3 {
        return Ok(CellValue::Error(
            "FLOOR.MATH expects 1 to 3 arguments (number, [significance], [mode])".to_string(),
        ));
    }

    let num_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let number = match to_number(&num_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "FLOOR.MATH number is not numeric".to_string(),
            ));
        },
    };

    let significance = if args.len() >= 2 {
        let sig_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&sig_val) {
            Some(n) if n.abs() >= EPS => n.abs(),
            Some(_) => {
                return Ok(CellValue::Error(
                    "FLOOR.MATH significance must be non-zero numeric".to_string(),
                ));
            },
            None => {
                return Ok(CellValue::Error(
                    "FLOOR.MATH significance is not numeric".to_string(),
                ));
            },
        }
    } else {
        1.0
    };

    let round_toward_zero = if args.len() == 3 {
        let mode_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match to_number(&mode_val) {
            Some(n) => n != 0.0,
            None => {
                return Ok(CellValue::Error(
                    "FLOOR.MATH mode is not numeric".to_string(),
                ));
            },
        }
    } else {
        false
    };

    if number == 0.0 {
        return Ok(CellValue::Int(0));
    }

    let quotient = number / significance;
    let rounded = if number >= 0.0 {
        quotient.floor()
    } else if round_toward_zero {
        quotient.ceil()
    } else {
        quotient.floor()
    };
    let result = rounded * significance;
    Ok(CellValue::Float(result))
}

pub(crate) async fn eval_floor_precise(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "FLOOR.PRECISE expects 1 or 2 arguments (number, [significance])".to_string(),
        ));
    }

    let num_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let number = match to_number(&num_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "FLOOR.PRECISE number is not numeric".to_string(),
            ));
        },
    };

    let significance = if args.len() == 2 {
        let sig_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&sig_val) {
            Some(n) if n.abs() >= EPS => n.abs(),
            Some(_) => {
                return Ok(CellValue::Error(
                    "FLOOR.PRECISE significance must be non-zero numeric".to_string(),
                ));
            },
            None => {
                return Ok(CellValue::Error(
                    "FLOOR.PRECISE significance is not numeric".to_string(),
                ));
            },
        }
    } else {
        1.0
    };

    if number == 0.0 {
        return Ok(CellValue::Int(0));
    }

    let quotient = number / significance;
    let rounded = quotient.floor();
    let result = rounded * significance;
    Ok(CellValue::Float(result))
}

pub(crate) async fn eval_ceiling_math(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 3 {
        return Ok(CellValue::Error(
            "CEILING.MATH expects 1 to 3 arguments (number, [significance], [mode])".to_string(),
        ));
    }

    let num_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let number = match to_number(&num_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "CEILING.MATH number is not numeric".to_string(),
            ));
        },
    };

    let significance = if args.len() >= 2 {
        let sig_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&sig_val) {
            Some(n) if n.abs() >= EPS => n,
            Some(_) => {
                return Ok(CellValue::Error(
                    "CEILING.MATH significance must be non-zero numeric".to_string(),
                ));
            },
            None => {
                return Ok(CellValue::Error(
                    "CEILING.MATH significance is not numeric".to_string(),
                ));
            },
        }
    } else if number >= 0.0 {
        1.0
    } else {
        -1.0
    };
    let significance = significance.abs();

    let round_away_from_zero = if args.len() == 3 {
        let mode_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match to_number(&mode_val) {
            Some(n) => n != 0.0,
            None => {
                return Ok(CellValue::Error(
                    "CEILING.MATH mode is not numeric".to_string(),
                ));
            },
        }
    } else {
        false
    };

    if number == 0.0 {
        return Ok(CellValue::Int(0));
    }

    let quotient = number / significance;
    let rounded = if number >= 0.0 {
        quotient.ceil()
    } else if round_away_from_zero {
        quotient.floor()
    } else {
        quotient.ceil()
    };
    let result = rounded * significance;
    Ok(CellValue::Float(result))
}

pub(crate) async fn eval_ceiling_precise(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "CEILING.PRECISE expects 1 or 2 arguments (number, [significance])".to_string(),
        ));
    }

    let num_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let number = match to_number(&num_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "CEILING.PRECISE number is not numeric".to_string(),
            ));
        },
    };

    let significance = if args.len() == 2 {
        let sig_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&sig_val) {
            Some(n) if n.abs() >= EPS => n.abs(),
            Some(_) => {
                return Ok(CellValue::Error(
                    "CEILING.PRECISE significance must be non-zero numeric".to_string(),
                ));
            },
            None => {
                return Ok(CellValue::Error(
                    "CEILING.PRECISE significance is not numeric".to_string(),
                ));
            },
        }
    } else {
        1.0
    };

    let quotient = number / significance;
    let rounded = quotient.ceil();
    let result = rounded * significance;
    Ok(CellValue::Float(result))
}

pub(crate) async fn eval_mod(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "MOD expects 2 arguments (number, divisor)".to_string(),
        ));
    }
    let number_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let divisor_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let number = match to_number(&number_val) {
        Some(n) => n,
        None => return Ok(CellValue::Error("MOD number is not numeric".to_string())),
    };
    let divisor = match to_number(&divisor_val) {
        Some(n) => n,
        None => return Ok(CellValue::Error("MOD divisor is not numeric".to_string())),
    };
    if divisor.abs() < EPS {
        return Ok(CellValue::Error("MOD divisor cannot be zero".to_string()));
    }
    let quotient = (number / divisor).floor();
    let result = number - divisor * quotient;
    Ok(number_result(result))
}

pub(crate) async fn eval_mround(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "MROUND expects 2 arguments (number, multiple)".to_string(),
        ));
    }
    let number_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let multiple_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let number = match to_number(&number_val) {
        Some(n) => n,
        None => return Ok(CellValue::Error("MROUND number is not numeric".to_string())),
    };
    let multiple = match to_number(&multiple_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "MROUND multiple is not numeric".to_string(),
            ));
        },
    };
    if multiple.abs() < EPS {
        return Ok(CellValue::Error(
            "MROUND multiple cannot be zero".to_string(),
        ));
    }
    if number.is_sign_positive() != multiple.is_sign_positive() && number != 0.0 {
        return Ok(CellValue::Error(
            "MROUND requires number and multiple to share the same sign".to_string(),
        ));
    }
    let quotient = number / multiple;
    let rounded = if quotient == 0.0 {
        0.0
    } else {
        round_away_from_zero(quotient)
    };
    let result = rounded * multiple;
    Ok(number_result(result))
}

pub(crate) async fn eval_quotient(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "QUOTIENT expects 2 arguments (numerator, denominator)".to_string(),
        ));
    }
    let numerator_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let denominator_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let numerator = match to_number(&numerator_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "QUOTIENT numerator is not numeric".to_string(),
            ));
        },
    };
    let denominator = match to_number(&denominator_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "QUOTIENT denominator is not numeric".to_string(),
            ));
        },
    };
    if denominator.abs() < EPS {
        return Ok(CellValue::Error(
            "QUOTIENT denominator cannot be zero".to_string(),
        ));
    }
    let value = (numerator / denominator).trunc();
    Ok(number_result(value))
}

pub(crate) async fn eval_trunc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "TRUNC expects 1 or 2 arguments (number, [num_digits])".to_string(),
        ));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => return Ok(CellValue::Error("TRUNC number is not numeric".to_string())),
    };
    let digits = if args.len() == 2 {
        let d_expr = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&d_expr) {
            Some(n) => n.trunc() as i32,
            None => {
                return Ok(CellValue::Error(
                    "TRUNC num_digits must be numeric".to_string(),
                ));
            },
        }
    } else {
        0
    };
    let factor = 10f64.powi(digits.abs());
    let result = if digits >= 0 {
        (x * factor).trunc() / factor
    } else {
        let d = digits.abs();
        (x / 10f64.powi(d)).trunc() * 10f64.powi(d)
    };
    Ok(number_result(result))
}

pub(crate) async fn eval_even(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("EVEN expects 1 argument".to_string()));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => return Ok(CellValue::Error("EVEN on non-numeric value".to_string())),
    };
    if x == 0.0 {
        return Ok(CellValue::Int(0));
    }
    let rounded = round_away_from_zero(x);
    let even = if is_even(rounded) {
        rounded
    } else if rounded > 0.0 {
        rounded + 1.0
    } else {
        rounded - 1.0
    };
    Ok(number_result(even))
}

pub(crate) async fn eval_odd(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ODD expects 1 argument".to_string()));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => return Ok(CellValue::Error("ODD on non-numeric value".to_string())),
    };
    if x == 0.0 {
        return Ok(CellValue::Int(1));
    }
    let rounded = round_away_from_zero(x);
    let odd = if is_even(rounded) {
        if rounded > 0.0 {
            rounded + 1.0
        } else {
            rounded - 1.0
        }
    } else {
        rounded
    };
    Ok(number_result(odd))
}

pub(crate) async fn eval_sign(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("SIGN expects 1 argument".to_string()));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = match to_number(&num) {
        Some(n) => n,
        None => return Ok(CellValue::Error("SIGN on non-numeric value".to_string())),
    };
    let value = if x > 0.0 {
        1
    } else if x < 0.0 {
        -1
    } else {
        0
    };
    Ok(CellValue::Int(value))
}

pub(crate) async fn eval_iso_ceiling(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "ISO.CEILING expects 1 or 2 arguments (number, [significance])".to_string(),
        ));
    }
    let num_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = match to_number(&num_val) {
        Some(n) => n,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let sig = if args.len() == 2 {
        let sig_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&sig_val) {
            Some(n) if n != 0.0 => n,
            _ => return Ok(CellValue::Error("#NUM!".to_string())),
        }
    } else {
        1.0
    };

    // ISO.CEILING rounds up to the nearest multiple of significance, regardless of the sign of the number.
    // However, if significance is negative, it uses the absolute value.
    let sig = sig.abs();
    let result = (x / sig).ceil() * sig;
    Ok(CellValue::Float(result))
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
    async fn test_eval_round_positive() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.14159), num_expr(2.0)];
        let result = eval_round(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.14).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_round_negative_digits() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1234.0), num_expr(-2.0)];
        let result = eval_round(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1200.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_round_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.14)];
        let result = eval_round(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 arguments")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_rounddown() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.99), num_expr(0.0)];
        let result = eval_rounddown(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_rounddown_negative() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-3.99), num_expr(0.0)];
        let result = eval_rounddown(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - (-3.0)).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_roundup() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.01), num_expr(0.0)];
        let result = eval_roundup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 4.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_roundup_negative() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-3.01), num_expr(0.0)];
        let result = eval_roundup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - (-4.0)).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_floor() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.7), num_expr(1.0)];
        let result = eval_floor(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_floor_default_sig() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.7)];
        let result = eval_floor(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_floor_zero_sig() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.7), num_expr(0.0)];
        let result = eval_floor(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("non-zero")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_ceiling() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.2), num_expr(1.0)];
        let result = eval_ceiling(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 4.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_ceiling_default_sig() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.2)];
        let result = eval_ceiling(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 4.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_ceiling_zero_sig() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.2), num_expr(0.0)];
        let result = eval_ceiling(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("non-zero")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_floor_math() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.7)];
        let result = eval_floor_math(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_floor_math_negative() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-3.7)];
        let result = eval_floor_math(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - (-4.0)).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_floor_math_mode() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-3.7), num_expr(1.0), num_expr(1.0)];
        let result = eval_floor_math(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - (-3.0)).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_floor_precise() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.7)];
        let result = eval_floor_precise(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_ceiling_math() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.2)];
        let result = eval_ceiling_math(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 4.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_ceiling_math_negative() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-3.2)];
        let result = eval_ceiling_math(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - (-3.0)).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_ceiling_precise() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.2)];
        let result = eval_ceiling_precise(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 4.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_mod() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(10.0), num_expr(3.0)];
        let result = eval_mod(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 1),
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected numeric result"),
        }
    }

    #[tokio::test]
    async fn test_eval_mod_negative() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-10.0), num_expr(3.0)];
        let result = eval_mod(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // Implementation returns 2: -10 = 3 * (-4) + 2
            // Result has same sign as divisor (positive)
            CellValue::Float(v) => assert!((v - 2.0).abs() < 1e-9),
            CellValue::Int(v) => assert_eq!(v, 2),
            _ => panic!("Expected numeric result"),
        }
    }

    #[tokio::test]
    async fn test_eval_mod_zero_divisor() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(10.0), num_expr(0.0)];
        let result = eval_mod(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("cannot be zero")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_mround() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(10.0), num_expr(3.0)];
        let result = eval_mround(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // Implementation rounds 10 to 12 for multiple 3 (rounds away from zero)
            CellValue::Int(v) => assert_eq!(v, 12),
            CellValue::Float(v) => assert!((v - 12.0).abs() < 1e-9),
            _ => panic!("Expected numeric result"),
        }
    }

    #[tokio::test]
    async fn test_eval_mround_different_signs() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(10.0), num_expr(-3.0)];
        let result = eval_mround(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("same sign")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_quotient() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(10.0), num_expr(3.0)];
        let result = eval_quotient(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 3),
            CellValue::Float(v) => assert!((v - 3.0).abs() < 1e-9),
            _ => panic!("Expected numeric result"),
        }
    }

    #[tokio::test]
    async fn test_eval_quotient_negative() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-10.0), num_expr(3.0)];
        let result = eval_quotient(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, -3),
            CellValue::Float(v) => assert!((v - (-3.0)).abs() < 1e-9),
            _ => panic!("Expected numeric result"),
        }
    }

    #[tokio::test]
    async fn test_eval_trunc() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.99)];
        let result = eval_trunc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 3),
            CellValue::Float(v) => assert!((v - 3.0).abs() < 1e-9),
            _ => panic!("Expected numeric result"),
        }
    }

    #[tokio::test]
    async fn test_eval_trunc_with_digits() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.14159), num_expr(2.0)];
        let result = eval_trunc(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.14).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_even_positive() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.2)];
        let result = eval_even(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 4),
            CellValue::Float(v) => assert!((v - 4.0).abs() < 1e-9),
            _ => panic!("Expected numeric result"),
        }
    }

    #[tokio::test]
    async fn test_eval_even_negative() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-3.2)];
        let result = eval_even(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, -4),
            CellValue::Float(v) => assert!((v - (-4.0)).abs() < 1e-9),
            _ => panic!("Expected numeric result"),
        }
    }

    #[tokio::test]
    async fn test_eval_even_zero() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_even(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 0),
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_odd_positive() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.2)];
        let result = eval_odd(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // ODD rounds away from zero, so 3.2 becomes 5 (the next odd number away from zero)
            CellValue::Int(v) => assert_eq!(v, 5),
            CellValue::Float(v) => assert!((v - 5.0).abs() < 1e-9),
            _ => panic!("Expected numeric result"),
        }
    }

    #[tokio::test]
    async fn test_eval_odd_negative() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-3.2)];
        let result = eval_odd(ctx, "Sheet1", &args).await.unwrap();
        match result {
            // ODD rounds away from zero, so -3.2 becomes -5
            CellValue::Int(v) => assert_eq!(v, -5),
            CellValue::Float(v) => assert!((v - (-5.0)).abs() < 1e-9),
            _ => panic!("Expected numeric result"),
        }
    }

    #[tokio::test]
    async fn test_eval_odd_zero() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_odd(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 1),
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sign_positive() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_sign(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 1),
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sign_negative() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-42.0)];
        let result = eval_sign(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, -1),
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sign_zero() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_sign(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 0),
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_iso_ceiling() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.2)];
        let result = eval_iso_ceiling(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 4.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }
}
