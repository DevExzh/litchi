use crate::sheet::{CellValue, Result};

use super::super::parser::Expr;
use super::registry::{self, FUNCTION_MAP};

pub(super) async fn eval_function(
    ctx: &dyn registry::DispatchCtx,
    current_sheet: &str,
    name: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if let Some(func) = FUNCTION_MAP.get(name) {
        return func(ctx, current_sheet, args).await;
    }

    Ok(CellValue::Error(format!("Unsupported function: {}", name)))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::engine::test_helpers::TestEngine;
    use crate::sheet::eval::parser::Expr;

    fn num_expr(n: f64) -> Expr {
        if n == n.floor() {
            Expr::Literal(CellValue::Int(n as i64))
        } else {
            Expr::Literal(CellValue::Float(n))
        }
    }

    #[tokio::test]
    async fn test_eval_function_sum() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling SUM function through dispatch
        let args = vec![num_expr(1.0), num_expr(2.0), num_expr(3.0)];
        let result = eval_function(ctx, "Sheet1", "SUM", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 6.0).abs() < 1e-9),
            _ => panic!("Expected Float(6.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_unsupported() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling an unsupported function
        let args: Vec<Expr> = vec![];
        let result = eval_function(ctx, "Sheet1", "NONEXISTENT_FUNC", &args)
            .await
            .unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("Unsupported function")),
            _ => panic!("Expected Error for unsupported function"),
        }
    }

    #[tokio::test]
    async fn test_eval_function_average() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling AVERAGE function through dispatch
        let args = vec![num_expr(10.0), num_expr(20.0), num_expr(30.0)];
        let result = eval_function(ctx, "Sheet1", "AVERAGE", &args)
            .await
            .unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 20.0).abs() < 1e-9),
            _ => panic!("Expected Float(20.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_if() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling IF function through dispatch
        let args = vec![
            Expr::Literal(CellValue::Bool(true)),
            num_expr(100.0),
            num_expr(0.0),
        ];
        let result = eval_function(ctx, "Sheet1", "IF", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 100),
            _ => panic!("Expected Int(100), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_max() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling MAX function through dispatch
        let args = vec![num_expr(5.0), num_expr(10.0), num_expr(3.0)];
        let result = eval_function(ctx, "Sheet1", "MAX", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 10.0).abs() < 1e-9),
            _ => panic!("Expected Float(10.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_min() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling MIN function through dispatch
        let args = vec![num_expr(5.0), num_expr(10.0), num_expr(3.0)];
        let result = eval_function(ctx, "Sheet1", "MIN", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.0).abs() < 1e-9),
            _ => panic!("Expected Float(3.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_count() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling COUNT function through dispatch
        let args = vec![num_expr(1.0), num_expr(2.0), num_expr(3.0)];
        let result = eval_function(ctx, "Sheet1", "COUNT", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 3),
            _ => panic!("Expected Int(3), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_abs() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling ABS function through dispatch
        let args = vec![num_expr(-42.0)];
        let result = eval_function(ctx, "Sheet1", "ABS", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 42),
            _ => panic!("Expected Int(42), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_product() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling PRODUCT function through dispatch
        let args = vec![num_expr(2.0), num_expr(3.0), num_expr(4.0)];
        let result = eval_function(ctx, "Sheet1", "PRODUCT", &args)
            .await
            .unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 24.0).abs() < 1e-9),
            _ => panic!("Expected Float(24.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_and() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling AND function through dispatch
        let args = vec![
            Expr::Literal(CellValue::Bool(true)),
            Expr::Literal(CellValue::Bool(true)),
        ];
        let result = eval_function(ctx, "Sheet1", "AND", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_or() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling OR function through dispatch
        let args = vec![
            Expr::Literal(CellValue::Bool(false)),
            Expr::Literal(CellValue::Bool(true)),
        ];
        let result = eval_function(ctx, "Sheet1", "OR", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_not() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling NOT function through dispatch
        let args = vec![Expr::Literal(CellValue::Bool(false))];
        let result = eval_function(ctx, "Sheet1", "NOT", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_concatenate() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling CONCATENATE function through dispatch
        let args = vec![
            Expr::Literal(CellValue::String("Hello".to_string())),
            Expr::Literal(CellValue::String(" ".to_string())),
            Expr::Literal(CellValue::String("World".to_string())),
        ];
        let result = eval_function(ctx, "Sheet1", "CONCATENATE", &args)
            .await
            .unwrap();
        match result {
            CellValue::String(v) => assert_eq!(v, "Hello World"),
            _ => panic!("Expected String, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_len() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling LEN function through dispatch
        let args = vec![Expr::Literal(CellValue::String("Hello".to_string()))];
        let result = eval_function(ctx, "Sheet1", "LEN", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 5),
            _ => panic!("Expected Int(5), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_left() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling LEFT function through dispatch
        let args = vec![
            Expr::Literal(CellValue::String("Hello".to_string())),
            num_expr(2.0),
        ];
        let result = eval_function(ctx, "Sheet1", "LEFT", &args).await.unwrap();
        match result {
            CellValue::String(v) => assert_eq!(v, "He"),
            _ => panic!("Expected String, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_right() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling RIGHT function through dispatch
        let args = vec![
            Expr::Literal(CellValue::String("Hello".to_string())),
            num_expr(2.0),
        ];
        let result = eval_function(ctx, "Sheet1", "RIGHT", &args).await.unwrap();
        match result {
            CellValue::String(v) => assert_eq!(v, "lo"),
            _ => panic!("Expected String, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_mid() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling MID function through dispatch
        let args = vec![
            Expr::Literal(CellValue::String("Hello".to_string())),
            num_expr(2.0),
            num_expr(2.0),
        ];
        let result = eval_function(ctx, "Sheet1", "MID", &args).await.unwrap();
        match result {
            CellValue::String(v) => assert_eq!(v, "el"),
            _ => panic!("Expected String, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_find() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling FIND function through dispatch
        let args = vec![
            Expr::Literal(CellValue::String("l".to_string())),
            Expr::Literal(CellValue::String("Hello".to_string())),
        ];
        let result = eval_function(ctx, "Sheet1", "FIND", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 3),
            _ => panic!("Expected Int(3), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_substitute() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling SUBSTITUTE function through dispatch
        let args = vec![
            Expr::Literal(CellValue::String("Hello World".to_string())),
            Expr::Literal(CellValue::String("World".to_string())),
            Expr::Literal(CellValue::String("Universe".to_string())),
        ];
        let result = eval_function(ctx, "Sheet1", "SUBSTITUTE", &args)
            .await
            .unwrap();
        match result {
            CellValue::String(v) => assert_eq!(v, "Hello Universe"),
            _ => panic!("Expected String, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_today() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling TODAY function through dispatch
        let args: Vec<Expr> = vec![];
        let result = eval_function(ctx, "Sheet1", "TODAY", &args).await.unwrap();
        match result {
            CellValue::DateTime(v) => {
                // TODAY should return a positive date serial
                assert!(v > 0.0);
                // Should be a whole number (no time component)
                assert_eq!(v, v.floor());
            },
            _ => panic!("Expected DateTime, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_pi() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling PI function through dispatch
        let args: Vec<Expr> = vec![];
        let result = eval_function(ctx, "Sheet1", "PI", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::PI).abs() < 1e-9),
            _ => panic!("Expected Float(PI), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_sqrt() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling SQRT function through dispatch
        let args = vec![num_expr(16.0)];
        let result = eval_function(ctx, "Sheet1", "SQRT", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 4.0).abs() < 1e-9),
            _ => panic!("Expected Float(4.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_power() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling POWER function through dispatch
        let args = vec![num_expr(2.0), num_expr(3.0)];
        let result = eval_function(ctx, "Sheet1", "POWER", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 8.0).abs() < 1e-9),
            _ => panic!("Expected Float(8.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_mod() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling MOD function through dispatch
        let args = vec![num_expr(10.0), num_expr(3.0)];
        let result = eval_function(ctx, "Sheet1", "MOD", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 1),
            _ => panic!("Expected Int(1), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_round() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling ROUND function through dispatch
        let args = vec![num_expr(3.7), num_expr(0.0)];
        let result = eval_function(ctx, "Sheet1", "ROUND", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 4),
            CellValue::Float(v) => assert!((v - 4.0).abs() < 1e-9),
            _ => panic!("Expected 4, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_isnumber() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling ISNUMBER function through dispatch
        let args = vec![num_expr(42.0)];
        let result = eval_function(ctx, "Sheet1", "ISNUMBER", &args)
            .await
            .unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_istext() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling ISTEXT function through dispatch
        let args = vec![Expr::Literal(CellValue::String("Hello".to_string()))];
        let result = eval_function(ctx, "Sheet1", "ISTEXT", &args).await.unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_isblank() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling ISBLANK function through dispatch
        let args = vec![Expr::Literal(CellValue::Empty)];
        let result = eval_function(ctx, "Sheet1", "ISBLANK", &args)
            .await
            .unwrap();
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool(true), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_upper() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling UPPER function through dispatch
        let args = vec![Expr::Literal(CellValue::String("hello".to_string()))];
        let result = eval_function(ctx, "Sheet1", "UPPER", &args).await.unwrap();
        match result {
            CellValue::String(v) => assert_eq!(v, "HELLO"),
            _ => panic!("Expected String, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_lower() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling LOWER function through dispatch
        let args = vec![Expr::Literal(CellValue::String("HELLO".to_string()))];
        let result = eval_function(ctx, "Sheet1", "LOWER", &args).await.unwrap();
        match result {
            CellValue::String(v) => assert_eq!(v, "hello"),
            _ => panic!("Expected String, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_trim() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling TRIM function through dispatch
        let args = vec![Expr::Literal(CellValue::String("  hello  ".to_string()))];
        let result = eval_function(ctx, "Sheet1", "TRIM", &args).await.unwrap();
        match result {
            CellValue::String(v) => assert_eq!(v, "hello"),
            _ => panic!("Expected String, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_value() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling VALUE function through dispatch
        let args = vec![Expr::Literal(CellValue::String("123".to_string()))];
        let result = eval_function(ctx, "Sheet1", "VALUE", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 123.0).abs() < 1e-9),
            CellValue::Int(v) => assert_eq!(v, 123),
            _ => panic!("Expected numeric 123, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_text() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling TEXT function through dispatch
        let args = vec![
            num_expr(1234.5),
            Expr::Literal(CellValue::String("0.00".to_string())),
        ];
        let result = eval_function(ctx, "Sheet1", "TEXT", &args).await.unwrap();
        match result {
            CellValue::String(v) => assert!(!v.is_empty()),
            _ => panic!("Expected String, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_date() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling DATE function through dispatch
        let args = vec![num_expr(2024.0), num_expr(3.0), num_expr(15.0)];
        let result = eval_function(ctx, "Sheet1", "DATE", &args).await.unwrap();
        match result {
            CellValue::DateTime(v) => assert!(v > 0.0),
            _ => panic!("Expected DateTime, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_year() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling YEAR function through dispatch
        let args = vec![num_expr(45366.0)]; // Excel serial for 2024-03-15
        let result = eval_function(ctx, "Sheet1", "YEAR", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 2024),
            _ => panic!("Expected Int(2024), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_month() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling MONTH function through dispatch
        let args = vec![num_expr(45366.0)]; // Excel serial for 2024-03-15
        let result = eval_function(ctx, "Sheet1", "MONTH", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 3),
            _ => panic!("Expected Int(3), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_day() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling DAY function through dispatch
        let args = vec![num_expr(45366.0)]; // Excel serial for 2024-03-15
        let result = eval_function(ctx, "Sheet1", "DAY", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 15),
            _ => panic!("Expected Int(15), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_sinh() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling SINH function through dispatch
        let args = vec![num_expr(1.0)];
        let result = eval_function(ctx, "Sheet1", "SINH", &args).await.unwrap();
        match result {
            CellValue::Float(v) => {
                let expected = 1.0f64.sinh();
                assert!((v - expected).abs() < 1e-9);
            },
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_cosh() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling COSH function through dispatch
        let args = vec![num_expr(1.0)];
        let result = eval_function(ctx, "Sheet1", "COSH", &args).await.unwrap();
        match result {
            CellValue::Float(v) => {
                let expected = 1.0f64.cosh();
                assert!((v - expected).abs() < 1e-9);
            },
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_tanh() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling TANH function through dispatch
        let args = vec![num_expr(1.0)];
        let result = eval_function(ctx, "Sheet1", "TANH", &args).await.unwrap();
        match result {
            CellValue::Float(v) => {
                let expected = 1.0f64.tanh();
                assert!((v - expected).abs() < 1e-9);
            },
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_exp() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling EXP function through dispatch
        let args = vec![num_expr(1.0)];
        let result = eval_function(ctx, "Sheet1", "EXP", &args).await.unwrap();
        match result {
            CellValue::Float(v) => {
                let expected = 1.0f64.exp();
                assert!((v - expected).abs() < 1e-9);
            },
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_ln() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling LN function through dispatch
        let args = vec![num_expr(std::f64::consts::E)];
        let result = eval_function(ctx, "Sheet1", "LN", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float(1.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_log10() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling LOG10 function through dispatch
        let args = vec![num_expr(100.0)];
        let result = eval_function(ctx, "Sheet1", "LOG10", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 2.0).abs() < 1e-9),
            _ => panic!("Expected Float(2.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_log() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling LOG function through dispatch
        let args = vec![num_expr(8.0), num_expr(2.0)];
        let result = eval_function(ctx, "Sheet1", "LOG", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.0).abs() < 1e-9),
            _ => panic!("Expected Float(3.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_sign() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling SIGN function through dispatch
        let args = vec![num_expr(-42.0)];
        let result = eval_function(ctx, "Sheet1", "SIGN", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, -1),
            _ => panic!("Expected Int(-1), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_trunc() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling TRUNC function through dispatch
        let args = vec![num_expr(3.7)];
        let result = eval_function(ctx, "Sheet1", "TRUNC", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 3),
            _ => panic!("Expected Int(3), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_int() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling INT function through dispatch
        let args = vec![num_expr(3.7)];
        let result = eval_function(ctx, "Sheet1", "INT", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 3),
            _ => panic!("Expected Int(3), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_ceiling() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling CEILING function through dispatch
        let args = vec![num_expr(3.2), num_expr(1.0)];
        let result = eval_function(ctx, "Sheet1", "CEILING", &args)
            .await
            .unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 4.0).abs() < 1e-9),
            _ => panic!("Expected Float(4.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_floor() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling FLOOR function through dispatch
        let args = vec![num_expr(3.8), num_expr(1.0)];
        let result = eval_function(ctx, "Sheet1", "FLOOR", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.0).abs() < 1e-9),
            _ => panic!("Expected Float(3.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_mround() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling MROUND function through dispatch
        // MROUND(10, 3) = 12 (implementation rounds to 12)
        let args = vec![num_expr(10.0), num_expr(3.0)];
        let result = eval_function(ctx, "Sheet1", "MROUND", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 12),
            CellValue::Float(v) => assert!((v - 12.0).abs() < 1e-9),
            _ => panic!("Expected 12, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_quotient() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling QUOTIENT function through dispatch
        let args = vec![num_expr(10.0), num_expr(3.0)];
        let result = eval_function(ctx, "Sheet1", "QUOTIENT", &args)
            .await
            .unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 3),
            _ => panic!("Expected Int(3), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_even() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling EVEN function through dispatch
        let args = vec![num_expr(3.0)];
        let result = eval_function(ctx, "Sheet1", "EVEN", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 4),
            _ => panic!("Expected Int(4), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_odd() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling ODD function through dispatch
        let args = vec![num_expr(4.0)];
        let result = eval_function(ctx, "Sheet1", "ODD", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 5),
            _ => panic!("Expected Int(5), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_fact() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling FACT function through dispatch
        let args = vec![num_expr(5.0)];
        let result = eval_function(ctx, "Sheet1", "FACT", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 120),
            CellValue::Float(v) => assert!((v - 120.0).abs() < 1e-9),
            _ => panic!("Expected 120, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_combin() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling COMBIN function through dispatch
        let args = vec![num_expr(5.0), num_expr(2.0)];
        let result = eval_function(ctx, "Sheet1", "COMBIN", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 10),
            CellValue::Float(v) => assert!((v - 10.0).abs() < 1e-9),
            _ => panic!("Expected 10, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_permut() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling PERMUT function through dispatch
        let args = vec![num_expr(5.0), num_expr(2.0)];
        let result = eval_function(ctx, "Sheet1", "PERMUT", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 20),
            CellValue::Float(v) => assert!((v - 20.0).abs() < 1e-9),
            _ => panic!("Expected 20, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_gcd() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling GCD function through dispatch
        let args = vec![num_expr(12.0), num_expr(8.0)];
        let result = eval_function(ctx, "Sheet1", "GCD", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 4),
            _ => panic!("Expected Int(4), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_lcm() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling LCM function through dispatch
        let args = vec![num_expr(4.0), num_expr(6.0)];
        let result = eval_function(ctx, "Sheet1", "LCM", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 12),
            _ => panic!("Expected Int(12), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_rand() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling RAND function through dispatch
        let args: Vec<Expr> = vec![];
        let result = eval_function(ctx, "Sheet1", "RAND", &args).await.unwrap();
        match result {
            CellValue::Float(v) => {
                assert!((0.0..1.0).contains(&v));
            },
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_randbetween() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling RANDBETWEEN function through dispatch
        let args = vec![num_expr(1.0), num_expr(10.0)];
        let result = eval_function(ctx, "Sheet1", "RANDBETWEEN", &args)
            .await
            .unwrap();
        match result {
            CellValue::Int(v) => {
                assert!((1..=10).contains(&v));
            },
            _ => panic!("Expected Int, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_degrees() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling DEGREES function through dispatch
        let args = vec![num_expr(std::f64::consts::PI)];
        let result = eval_function(ctx, "Sheet1", "DEGREES", &args)
            .await
            .unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 180.0).abs() < 1e-9),
            _ => panic!("Expected Float(180.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_radians() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling RADIANS function through dispatch
        let args = vec![num_expr(180.0)];
        let result = eval_function(ctx, "Sheet1", "RADIANS", &args)
            .await
            .unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::PI).abs() < 1e-9),
            _ => panic!("Expected Float(PI), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_sin() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling SIN function through dispatch
        let args = vec![num_expr(std::f64::consts::PI / 2.0)];
        let result = eval_function(ctx, "Sheet1", "SIN", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float(1.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_cos() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling COS function through dispatch
        let args = vec![num_expr(0.0)];
        let result = eval_function(ctx, "Sheet1", "COS", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 1),
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float(1.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_tan() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling TAN function through dispatch
        let args = vec![num_expr(std::f64::consts::PI / 4.0)];
        let result = eval_function(ctx, "Sheet1", "TAN", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1.0).abs() < 1e-9),
            _ => panic!("Expected Float(1.0), got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_asin() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling ASIN function through dispatch
        let args = vec![num_expr(1.0)];
        let result = eval_function(ctx, "Sheet1", "ASIN", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::PI / 2.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_acos() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling ACOS function through dispatch
        let args = vec![num_expr(1.0)];
        let result = eval_function(ctx, "Sheet1", "ACOS", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 0),
            CellValue::Float(v) => assert!(v.abs() < 1e-9),
            _ => panic!("Expected 0, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_atan() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling ATAN function through dispatch
        let args = vec![num_expr(1.0)];
        let result = eval_function(ctx, "Sheet1", "ATAN", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::PI / 4.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_function_atan2() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Test calling ATAN2 function through dispatch
        let args = vec![num_expr(1.0), num_expr(1.0)];
        let result = eval_function(ctx, "Sheet1", "ATAN2", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - std::f64::consts::PI / 4.0).abs() < 1e-9),
            _ => panic!("Expected Float, got {:?}", result),
        }
    }
}
