use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, for_each_value_in_expr, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::helpers::{
    combination, double_factorial, factorial, number_result, permutation, to_int_if_whole,
};

pub(crate) async fn eval_gcd(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "GCD expects at least 1 argument".to_string(),
        ));
    }
    let mut result: Option<u128> = None;
    let mut invalid: Option<CellValue> = None;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |value| {
            if invalid.is_some() {
                return Ok(());
            }
            let number = match to_number(value) {
                Some(n) => n,
                None => return Ok(()),
            };
            let int_value = match to_int_if_whole(number.abs()) {
                Some(v) => v as u128,
                None => {
                    invalid = Some(CellValue::Error(
                        "GCD arguments must be integers".to_string(),
                    ));
                    return Ok(());
                },
            };
            result = Some(match result {
                Some(current) => gcd_u128(current, int_value),
                None => int_value,
            });
            Ok(())
        })
        .await?;
        if invalid.is_some() {
            break;
        }
    }
    if let Some(err) = invalid {
        return Ok(err);
    }
    match result {
        Some(value) => Ok(CellValue::Int(value as i64)),
        None => Ok(CellValue::Error(
            "GCD received no numeric arguments".to_string(),
        )),
    }
}

pub(crate) async fn eval_lcm(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "LCM expects at least 1 argument".to_string(),
        ));
    }
    let mut result: Option<u128> = None;
    let mut invalid: Option<CellValue> = None;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |value| {
            if invalid.is_some() {
                return Ok(());
            }
            let number = match to_number(value) {
                Some(n) => n,
                None => return Ok(()),
            };
            if number < 0.0 {
                invalid = Some(CellValue::Error("#NUM!".to_string()));
                return Ok(());
            }
            let int_value = match to_int_if_whole(number) {
                Some(v) => v as u128,
                None => {
                    invalid = Some(CellValue::Error(
                        "LCM arguments must be integers".to_string(),
                    ));
                    return Ok(());
                },
            };
            if int_value == 0 {
                result = Some(0);
                return Ok(());
            }
            result = Some(match result {
                Some(0) => 0,
                Some(current) => lcm_u128(current, int_value),
                None => int_value,
            });
            Ok(())
        })
        .await?;
        if invalid.is_some() {
            break;
        }
    }
    if let Some(err) = invalid {
        return Ok(err);
    }
    match result {
        Some(value) => {
            if value > i64::MAX as u128 {
                Ok(CellValue::Float(value as f64))
            } else {
                Ok(CellValue::Int(value as i64))
            }
        },
        None => Ok(CellValue::Error(
            "LCM received no numeric arguments".to_string(),
        )),
    }
}

pub(crate) async fn eval_fact(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("FACT expects 1 argument".to_string()));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let n = match to_number(&num) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => {
            return Ok(CellValue::Error(
                "FACT requires non-negative number".to_string(),
            ));
        },
        None => return Ok(CellValue::Error("FACT on non-numeric value".to_string())),
    };
    if n > 170.0 {
        return Ok(CellValue::Error("FACT input too large".to_string()));
    }
    Ok(number_result(factorial(n as u64)))
}

pub(crate) async fn eval_factdouble(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error(
            "FACTDOUBLE expects 1 argument".to_string(),
        ));
    }
    let num = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let n = match to_number(&num) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => {
            return Ok(CellValue::Error(
                "FACTDOUBLE requires non-negative number".to_string(),
            ));
        },
        None => {
            return Ok(CellValue::Error(
                "FACTDOUBLE on non-numeric value".to_string(),
            ));
        },
    };
    if n > 170.0 {
        return Ok(CellValue::Error("FACTDOUBLE input too large".to_string()));
    }
    Ok(number_result(double_factorial(n as u64)))
}

pub(crate) async fn eval_combin(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("COMBIN expects 2 arguments".to_string()));
    }
    let n_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let k_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let n = match to_number(&n_val) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => {
            return Ok(CellValue::Error(
                "COMBIN requires non-negative number".to_string(),
            ));
        },
        None => {
            return Ok(CellValue::Error(
                "COMBIN number must be numeric".to_string(),
            ));
        },
    };
    let k = match to_number(&k_val) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => {
            return Ok(CellValue::Error(
                "COMBIN requires non-negative number_chosen".to_string(),
            ));
        },
        None => {
            return Ok(CellValue::Error(
                "COMBIN number_chosen must be numeric".to_string(),
            ));
        },
    };
    if k > n {
        return Ok(CellValue::Error(
            "COMBIN requires number_chosen <= number".to_string(),
        ));
    }
    Ok(number_result(combination(n as u64, k as u64)))
}

pub(crate) async fn eval_combina(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("COMBINA expects 2 arguments".to_string()));
    }
    let n_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let k_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let n = match to_number(&n_val) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => {
            return Ok(CellValue::Error(
                "COMBINA requires non-negative number".to_string(),
            ));
        },
        None => {
            return Ok(CellValue::Error(
                "COMBINA number must be numeric".to_string(),
            ));
        },
    };
    let k = match to_number(&k_val) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => {
            return Ok(CellValue::Error(
                "COMBINA requires non-negative number_chosen".to_string(),
            ));
        },
        None => {
            return Ok(CellValue::Error(
                "COMBINA number_chosen must be numeric".to_string(),
            ));
        },
    };
    if n == 0.0 {
        return if k == 0.0 {
            Ok(CellValue::Int(1))
        } else {
            Ok(CellValue::Error(
                "COMBINA requires number > 0 when number_chosen > 0".to_string(),
            ))
        };
    }
    let total = (n as u64) + (k as u64) - 1;
    Ok(number_result(combination(total, k as u64)))
}

pub(crate) async fn eval_permut(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("PERMUT expects 2 arguments".to_string()));
    }
    let n_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let k_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let n = match to_number(&n_val) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => {
            return Ok(CellValue::Error(
                "PERMUT requires non-negative number".to_string(),
            ));
        },
        None => {
            return Ok(CellValue::Error(
                "PERMUT number must be numeric".to_string(),
            ));
        },
    };
    let k = match to_number(&k_val) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => {
            return Ok(CellValue::Error(
                "PERMUT requires non-negative number_chosen".to_string(),
            ));
        },
        None => {
            return Ok(CellValue::Error(
                "PERMUT number_chosen must be numeric".to_string(),
            ));
        },
    };
    if k > n {
        return Ok(CellValue::Error(
            "PERMUT requires number_chosen <= number".to_string(),
        ));
    }
    Ok(number_result(permutation(n as u64, k as u64)))
}

pub(crate) async fn eval_permutationa(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "PERMUTATIONA expects 2 arguments".to_string(),
        ));
    }
    let n_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let k_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let n = match to_number(&n_val) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => {
            return Ok(CellValue::Error(
                "PERMUTATIONA requires non-negative number".to_string(),
            ));
        },
        None => {
            return Ok(CellValue::Error(
                "PERMUTATIONA number must be numeric".to_string(),
            ));
        },
    };
    let k = match to_number(&k_val) {
        Some(v) if v >= 0.0 => v.trunc(),
        Some(_) => {
            return Ok(CellValue::Error(
                "PERMUTATIONA requires non-negative number_chosen".to_string(),
            ));
        },
        None => {
            return Ok(CellValue::Error(
                "PERMUTATIONA number_chosen must be numeric".to_string(),
            ));
        },
    };
    let base = n as f64;
    let mut result = 1.0;
    for _ in 0..(k as u64) {
        result *= base;
    }
    Ok(number_result(result))
}

pub(crate) async fn eval_multinomial(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "MULTINOMIAL expects at least 1 argument".to_string(),
        ));
    }
    let mut sum: u64 = 0;
    let mut denom = 1.0;
    for expr in args {
        let value = evaluate_expression(ctx, current_sheet, expr).await?;
        let number = match to_number(&value) {
            Some(v) if v >= 0.0 => v.trunc(),
            Some(_) => {
                return Ok(CellValue::Error(
                    "MULTINOMIAL requires non-negative arguments".to_string(),
                ));
            },
            None => {
                return Ok(CellValue::Error(
                    "MULTINOMIAL arguments must be numeric".to_string(),
                ));
            },
        };
        let component = number as u64;
        sum = match sum.checked_add(component) {
            Some(total) => total,
            None => {
                return Ok(CellValue::Error(
                    "MULTINOMIAL argument sum too large".to_string(),
                ));
            },
        };
        denom *= factorial(component);
    }
    if sum > 170 {
        return Ok(CellValue::Error(
            "MULTINOMIAL argument sum too large".to_string(),
        ));
    }
    let numerator = factorial(sum);
    Ok(number_result(numerator / denom))
}

fn gcd_u128(mut a: u128, mut b: u128) -> u128 {
    while b != 0 {
        let r = a % b;
        a = b;
        b = r;
    }
    a
}

fn lcm_u128(a: u128, b: u128) -> u128 {
    if a == 0 || b == 0 {
        return 0;
    }
    let g = gcd_u128(a, b);
    (a / g) * b
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
    async fn test_eval_fact() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(5.0)];
        let result = eval_fact(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(120)); // 5! = 120
    }

    #[tokio::test]
    async fn test_eval_fact_zero() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_fact(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(1)); // 0! = 1
    }

    #[tokio::test]
    async fn test_eval_fact_negative() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(-5.0)];
        let result = eval_fact(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("non-negative")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_combin() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // C(5, 2) = 10
        let args = vec![num_expr(5.0), num_expr(2.0)];
        let result = eval_combin(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(10));
    }

    #[tokio::test]
    async fn test_eval_combin_k_greater_n() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.0), num_expr(5.0)];
        let result = eval_combin(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("number_chosen <= number")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_permut() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // P(5, 2) = 20
        let args = vec![num_expr(5.0), num_expr(2.0)];
        let result = eval_permut(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(20));
    }

    #[tokio::test]
    async fn test_eval_gcd() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // GCD(48, 18) = 6
        let args = vec![num_expr(48.0), num_expr(18.0)];
        let result = eval_gcd(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(6));
    }

    #[tokio::test]
    async fn test_eval_gcd_single_arg() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_gcd(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(42));
    }

    #[tokio::test]
    async fn test_eval_lcm() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // LCM(4, 6) = 12
        let args = vec![num_expr(4.0), num_expr(6.0)];
        let result = eval_lcm(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(12));
    }

    #[tokio::test]
    async fn test_eval_factdouble() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // 6!! = 6 * 4 * 2 = 48
        let args = vec![num_expr(6.0)];
        let result = eval_factdouble(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(48));
    }

    #[tokio::test]
    async fn test_eval_multinomial() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // MULTINOMIAL(2, 3, 4) = (2+3+4)! / (2! * 3! * 4!) = 9! / (2! * 3! * 4!) = 1260
        let args = vec![num_expr(2.0), num_expr(3.0), num_expr(4.0)];
        let result = eval_multinomial(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(1260));
    }
}
