use crate::sheet::{CellValue, Result};

use super::helpers::{MAX_BITWISE_VALUE, bit_operand_value, bit_shift_value};
use crate::sheet::eval::engine::{EvalCtx, evaluate_expression};
use crate::sheet::eval::parser::Expr;

pub(crate) async fn eval_bitand(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("BITAND expects 2 arguments".to_string()));
    }
    let left_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = left_value {
        return Ok(left_value);
    }
    let right_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    if let CellValue::Error(_) = right_value {
        return Ok(right_value);
    }
    let left = match bit_operand_value(&left_value, "BITAND") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    let right = match bit_operand_value(&right_value, "BITAND") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    Ok(CellValue::Int((left & right) as i64))
}

pub(crate) async fn eval_bitor(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("BITOR expects 2 arguments".to_string()));
    }
    let left_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = left_value {
        return Ok(left_value);
    }
    let right_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    if let CellValue::Error(_) = right_value {
        return Ok(right_value);
    }
    let left = match bit_operand_value(&left_value, "BITOR") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    let right = match bit_operand_value(&right_value, "BITOR") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    Ok(CellValue::Int((left | right) as i64))
}

pub(crate) async fn eval_bitxor(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("BITXOR expects 2 arguments".to_string()));
    }
    let left_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = left_value {
        return Ok(left_value);
    }
    let right_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    if let CellValue::Error(_) = right_value {
        return Ok(right_value);
    }
    let left = match bit_operand_value(&left_value, "BITXOR") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    let right = match bit_operand_value(&right_value, "BITXOR") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    Ok(CellValue::Int((left ^ right) as i64))
}

pub(crate) async fn eval_bitlshift(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "BITLSHIFT expects 2 arguments".to_string(),
        ));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let shift_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    if let CellValue::Error(_) = shift_value {
        return Ok(shift_value);
    }
    let number = match bit_operand_value(&number_value, "BITLSHIFT") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    let shift = match bit_shift_value(&shift_value, "BITLSHIFT") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    let shifted = match number.checked_shl(shift) {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(
                "BITLSHIFT shift must be between 0 and 53".to_string(),
            ));
        },
    };
    if shifted > MAX_BITWISE_VALUE {
        return Ok(CellValue::Error(
            "BITLSHIFT result exceeds 48 bits".to_string(),
        ));
    }
    Ok(CellValue::Int(shifted as i64))
}

pub(crate) async fn eval_bitrshift(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "BITRSHIFT expects 2 arguments".to_string(),
        ));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let shift_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    if let CellValue::Error(_) = shift_value {
        return Ok(shift_value);
    }
    let number = match bit_operand_value(&number_value, "BITRSHIFT") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    let shift = match bit_shift_value(&shift_value, "BITRSHIFT") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    Ok(CellValue::Int((number >> shift) as i64))
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
    async fn test_eval_bitand() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // 6 = 110, 3 = 011, result = 010 = 2
        let args = vec![num_expr(6.0), num_expr(3.0)];
        let result = eval_bitand(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(2));
    }

    #[tokio::test]
    async fn test_eval_bitor() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // 6 = 110, 3 = 011, result = 111 = 7
        let args = vec![num_expr(6.0), num_expr(3.0)];
        let result = eval_bitor(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(7));
    }

    #[tokio::test]
    async fn test_eval_bitxor() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // 6 = 110, 3 = 011, result = 101 = 5
        let args = vec![num_expr(6.0), num_expr(3.0)];
        let result = eval_bitxor(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(5));
    }

    #[tokio::test]
    async fn test_eval_bitlshift() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // 5 << 2 = 20
        let args = vec![num_expr(5.0), num_expr(2.0)];
        let result = eval_bitlshift(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(20));
    }

    #[tokio::test]
    async fn test_eval_bitrshift() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // 20 >> 2 = 5
        let args = vec![num_expr(20.0), num_expr(2.0)];
        let result = eval_bitrshift(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(5));
    }

    #[tokio::test]
    async fn test_eval_bitand_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(6.0)];
        let result = eval_bitand(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 arguments")),
            _ => panic!("Expected Error result"),
        }
    }
}
