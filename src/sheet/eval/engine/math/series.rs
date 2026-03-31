use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, for_each_value_in_expr, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};
use rand::RngExt;

use super::helpers::flatten_numeric_values;

pub(crate) async fn eval_sumsq(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut total = 0.0f64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |value| {
            if let Some(n) = to_number(value) {
                total += n * n;
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_sumx2my2(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "SUMX2MY2 expects 2 arguments (array_x, array_y)".to_string(),
        ));
    }
    let x_vals = flatten_numeric_values(ctx, current_sheet, &args[0]).await?;
    let y_vals = flatten_numeric_values(ctx, current_sheet, &args[1]).await?;
    if x_vals.len() != y_vals.len() {
        return Ok(CellValue::Error(
            "SUMX2MY2 requires arrays of the same size".to_string(),
        ));
    }
    let total = x_vals
        .iter()
        .zip(y_vals.iter())
        .fold(0.0, |acc, (&x, &y)| acc + x * x - y * y);
    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_sumx2py2(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "SUMX2PY2 expects 2 arguments (array_x, array_y)".to_string(),
        ));
    }
    let x_vals = flatten_numeric_values(ctx, current_sheet, &args[0]).await?;
    let y_vals = flatten_numeric_values(ctx, current_sheet, &args[1]).await?;
    if x_vals.len() != y_vals.len() {
        return Ok(CellValue::Error(
            "SUMX2PY2 requires arrays of the same size".to_string(),
        ));
    }
    let total = x_vals
        .iter()
        .zip(y_vals.iter())
        .fold(0.0, |acc, (&x, &y)| acc + x * x + y * y);
    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_sumxmy2(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "SUMXMY2 expects 2 arguments (array_x, array_y)".to_string(),
        ));
    }
    let x_vals = flatten_numeric_values(ctx, current_sheet, &args[0]).await?;
    let y_vals = flatten_numeric_values(ctx, current_sheet, &args[1]).await?;
    if x_vals.len() != y_vals.len() {
        return Ok(CellValue::Error(
            "SUMXMY2 requires arrays of the same size".to_string(),
        ));
    }
    let total = x_vals.iter().zip(y_vals.iter()).fold(0.0, |acc, (&x, &y)| {
        let diff = x - y;
        acc + diff * diff
    });
    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_seriessum(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 4 {
        return Ok(CellValue::Error(
            "SERIESSUM expects 4 arguments (x, n, m, coefficients)".to_string(),
        ));
    }

    let x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let n = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let m = match to_number(&evaluate_expression(ctx, current_sheet, &args[2]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let mut coefficients = Vec::new();
    for_each_value_in_expr(ctx, current_sheet, &args[3], |val| {
        if let Some(c) = to_number(val) {
            coefficients.push(c);
        }
        Ok(())
    })
    .await?;

    if coefficients.is_empty() {
        return Ok(CellValue::Float(0.0));
    }

    let mut total = 0.0;
    for (i, &coeff) in coefficients.iter().enumerate() {
        let power = n + (i as f64) * m;
        total += coeff * x.powf(power);
    }

    Ok(CellValue::Float(total))
}

pub(crate) async fn eval_sequence(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 4 {
        return Ok(CellValue::Error(
            "SEQUENCE expects 1 to 4 arguments (rows, [columns], [start], [step])".to_string(),
        ));
    }

    let rows = match args.first() {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) if n >= 0.0 => n.trunc() as usize,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1,
    };

    let cols = match args.get(1) {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) if n >= 0.0 => n.trunc() as usize,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1,
    };

    let start = match args.get(2) {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) => n,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1.0,
    };

    let step = match args.get(3) {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) => n,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1.0,
    };

    if rows == 0 || cols == 0 {
        return Ok(CellValue::Error("#CALC!".to_string()));
    }

    let _step = step; // Avoid unused warning for now
    // Since CellValue doesn't support arrays yet, return the first value (start)
    Ok(CellValue::Float(start))
}

pub(crate) async fn eval_vstack(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "VSTACK expects at least 1 argument".to_string(),
        ));
    }
    // Placeholder: return first value of first argument
    evaluate_expression(ctx, current_sheet, &args[0]).await
}

pub(crate) async fn eval_hstack(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "HSTACK expects at least 1 argument".to_string(),
        ));
    }
    // Placeholder: return first value of first argument
    evaluate_expression(ctx, current_sheet, &args[0]).await
}

pub(crate) async fn eval_wrapcols(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "WRAPCOLS expects 2 or 3 arguments".to_string(),
        ));
    }
    // Placeholder: return first value of first argument
    evaluate_expression(ctx, current_sheet, &args[0]).await
}

pub(crate) async fn eval_wraprows(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "WRAPROWS expects 2 or 3 arguments".to_string(),
        ));
    }
    // Placeholder: return first value of first argument
    evaluate_expression(ctx, current_sheet, &args[0]).await
}

pub(crate) async fn eval_randarray(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() > 5 {
        return Ok(CellValue::Error(
            "RANDARRAY expects 0 to 5 arguments (rows, [columns], [min], [max], [whole_number])"
                .to_string(),
        ));
    }

    let rows = match args.first() {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) if n >= 0.0 => n.trunc() as usize,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1,
    };

    let cols = match args.get(1) {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) if n >= 0.0 => n.trunc() as usize,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1,
    };

    let min = match args.get(2) {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) => n,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 0.0,
    };

    let max = match args.get(3) {
        Some(expr) => match to_number(&evaluate_expression(ctx, current_sheet, expr).await?) {
            Some(n) => n,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        },
        None => 1.0,
    };

    let whole_number = match args.get(4) {
        Some(expr) => {
            let val = evaluate_expression(ctx, current_sheet, expr).await?;
            match val {
                CellValue::Bool(b) => b,
                _ => match to_number(&val) {
                    Some(n) => n != 0.0,
                    None => false,
                },
            }
        },
        None => false,
    };

    if rows == 0 || cols == 0 {
        return Ok(CellValue::Error("#CALC!".to_string()));
    }

    if min > max {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }

    let mut rng = rand::rng();
    let val = if whole_number {
        let bottom = min.ceil() as i64;
        let top = max.floor() as i64;
        if bottom > top {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }
        CellValue::Int(rng.random_range(bottom..=top))
    } else {
        CellValue::Float(min + (max - min) * rng.random::<f64>())
    };

    // Since CellValue doesn't support arrays yet, return a single random value
    Ok(val)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::engine::test_helpers::TestEngine;
    use crate::sheet::eval::parser::Expr;
    use crate::sheet::eval::parser::ast::RangeRef;

    fn num_expr(n: f64) -> Expr {
        if n == n.floor() {
            Expr::Literal(CellValue::Int(n as i64))
        } else {
            Expr::Literal(CellValue::Float(n))
        }
    }

    fn range_expr(sheet: &str, start_row: u32, start_col: u32, end_row: u32, end_col: u32) -> Expr {
        Expr::Range(RangeRef {
            sheet: sheet.to_string(),
            start_row,
            start_col,
            end_row,
            end_col,
        })
    }

    #[tokio::test]
    async fn test_eval_sumsq() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Add some values to the engine
        let values = vec![
            CellValue::Int(1),
            CellValue::Int(2),
            CellValue::Int(3),
            CellValue::Int(4),
        ];
        engine.add_range("Sheet1", 1, 1, 2, 2, values);

        // SUMSQ(Sheet1!A1:B2) = 1^2 + 2^2 + 3^2 + 4^2 = 30
        let args = vec![range_expr("Sheet1", 1, 1, 2, 2)];
        let result = eval_sumsq(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 30.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sumsq_single_values() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // SUMSQ(3, 4) = 3^2 + 4^2 = 25
        let args = vec![num_expr(3.0), num_expr(4.0)];
        let result = eval_sumsq(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 25.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sumx2my2() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Arrays: [1, 2, 3] and [1, 2, 3]
        // Sum of (x^2 - y^2) = (1-1) + (4-4) + (9-9) = 0
        let values1 = vec![CellValue::Int(1), CellValue::Int(2), CellValue::Int(3)];
        let values2 = vec![CellValue::Int(1), CellValue::Int(2), CellValue::Int(3)];
        engine.add_range("Sheet1", 1, 1, 1, 3, values1);
        engine.add_range("Sheet1", 2, 1, 1, 3, values2);

        let args = vec![
            range_expr("Sheet1", 1, 1, 1, 3),
            range_expr("Sheet1", 2, 1, 2, 3),
        ];
        let result = eval_sumx2my2(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sumx2py2() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Arrays: [3, 4] and [0, 0]
        // Sum of (x^2 + y^2) = (9+0) + (16+0) = 25
        let values1 = vec![CellValue::Int(3), CellValue::Int(4)];
        let values2 = vec![CellValue::Int(0), CellValue::Int(0)];
        engine.add_range("Sheet1", 1, 1, 1, 2, values1);
        engine.add_range("Sheet1", 2, 1, 1, 2, values2);

        let args = vec![
            range_expr("Sheet1", 1, 1, 1, 2),
            range_expr("Sheet1", 2, 1, 2, 2),
        ];
        let result = eval_sumx2py2(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 25.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sumxmy2() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Arrays: [1, 2, 3] and [0, 0, 0]
        // Sum of (x - y)^2 = 1 + 4 + 9 = 14
        let values1 = vec![CellValue::Int(1), CellValue::Int(2), CellValue::Int(3)];
        let values2 = vec![CellValue::Int(0), CellValue::Int(0), CellValue::Int(0)];
        engine.add_range("Sheet1", 1, 1, 1, 3, values1);
        engine.add_range("Sheet1", 2, 1, 1, 3, values2);

        let args = vec![
            range_expr("Sheet1", 1, 1, 1, 3),
            range_expr("Sheet1", 2, 1, 2, 3),
        ];
        let result = eval_sumxmy2(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 14.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sumxmy2_wrong_size() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let values1 = vec![CellValue::Int(1), CellValue::Int(2)];
        let values2 = vec![CellValue::Int(1)];
        engine.add_range("Sheet1", 1, 1, 1, 2, values1);
        engine.add_range("Sheet1", 2, 1, 1, 1, values2);

        let args = vec![
            range_expr("Sheet1", 1, 1, 1, 2),
            range_expr("Sheet1", 2, 1, 2, 1),
        ];
        let result = eval_sumxmy2(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("same size")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_seriessum() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // x=2, n=0, m=1, coefficients=[1, 2, 3]
        // = 1*2^0 + 2*2^1 + 3*2^2 = 1 + 4 + 12 = 17
        let coeffs = vec![CellValue::Int(1), CellValue::Int(2), CellValue::Int(3)];
        engine.add_range("Sheet1", 1, 1, 1, 3, coeffs);

        let args = vec![
            num_expr(2.0),
            num_expr(0.0),
            num_expr(1.0),
            range_expr("Sheet1", 1, 1, 1, 3),
        ];
        let result = eval_seriessum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 17.0).abs() < 1e-9),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_seriessum_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(2.0), num_expr(0.0)];
        let result = eval_seriessum(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 4 arguments")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sequence() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // SEQUENCE(3, 2) - returns first value as placeholder
        let args = vec![num_expr(3.0), num_expr(2.0)];
        let result = eval_sequence(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9), // Returns start value
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_sequence_zero_rows() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(0.0)];
        let result = eval_sequence(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#CALC!")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_vstack() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_vstack(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(42));
    }

    #[tokio::test]
    async fn test_eval_vstack_no_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_vstack(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects at least 1 argument")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_eval_hstack() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_hstack(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(42));
    }

    #[tokio::test]
    async fn test_eval_wrapcols() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // WRAPCOLS with literal value
        let args = vec![num_expr(42.0), num_expr(2.0)];
        let result = eval_wrapcols(ctx, "Sheet1", &args).await.unwrap();
        // Returns the value as placeholder
        assert_eq!(result, CellValue::Int(42));
    }

    #[tokio::test]
    async fn test_eval_wraprows() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // WRAPROWS with literal value
        let args = vec![num_expr(42.0), num_expr(2.0)];
        let result = eval_wraprows(ctx, "Sheet1", &args).await.unwrap();
        // Returns the value as placeholder
        assert_eq!(result, CellValue::Int(42));
    }

    #[tokio::test]
    async fn test_eval_randarray() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.0), num_expr(2.0)];
        let result = eval_randarray(ctx, "Sheet1", &args).await.unwrap();
        // Returns a single random value between 0 and 1
        match result {
            CellValue::Float(v) => assert!((0.0..=1.0).contains(&v)),
            _ => panic!("Expected Float result"),
        }
    }

    #[tokio::test]
    async fn test_eval_randarray_whole_number() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // RANDARRAY(3, 2, 1, 10, TRUE)
        let args = vec![
            num_expr(3.0),
            num_expr(2.0),
            num_expr(1.0),
            num_expr(10.0),
            num_expr(1.0), // TRUE
        ];
        let result = eval_randarray(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert!((1..=10).contains(&v)),
            _ => panic!("Expected Int result"),
        }
    }

    #[tokio::test]
    async fn test_eval_randarray_min_greater_max() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            num_expr(3.0),
            num_expr(2.0),
            num_expr(10.0),
            num_expr(1.0), // min > max
        ];
        let result = eval_randarray(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#VALUE!")),
            _ => panic!("Expected Error result"),
        }
    }
}
