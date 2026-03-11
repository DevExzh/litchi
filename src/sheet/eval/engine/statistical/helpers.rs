use std::cmp::Ordering;
use std::result::Result as StdResult;

use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

pub(super) const EPS: f64 = 1e-12;

pub(super) async fn number_arg(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    expr: &Expr,
) -> Result<Option<f64>> {
    let v = evaluate_expression(ctx, current_sheet, expr).await?;
    Ok(to_number(&v))
}

pub(super) fn collect_numeric_values(
    values: Vec<CellValue>,
    func_name: &str,
) -> StdResult<Vec<f64>, CellValue> {
    let mut numbers = Vec::new();
    for value in values {
        match value {
            CellValue::Error(msg) => return Err(CellValue::Error(msg)),
            other => {
                if let Some(n) = to_number(&other) {
                    if n.is_nan() {
                        return Err(CellValue::Error(format!(
                            "{func_name} encountered an invalid numeric value"
                        )));
                    }
                    numbers.push(n);
                }
            },
        }
    }

    if numbers.is_empty() {
        return Err(CellValue::Error(format!(
            "{func_name} requires at least one numeric value in the array"
        )));
    }

    numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    Ok(numbers)
}

pub(super) fn collect_numeric_values_unsorted(
    values: Vec<CellValue>,
    func_name: &str,
) -> StdResult<Vec<f64>, CellValue> {
    let mut numbers = Vec::new();
    for value in values {
        match value {
            CellValue::Error(msg) => return Err(CellValue::Error(msg)),
            other => {
                if let Some(n) = to_number(&other) {
                    if n.is_nan() {
                        return Err(CellValue::Error(format!(
                            "{func_name} encountered an invalid numeric value"
                        )));
                    }
                    numbers.push(n);
                }
            },
        }
    }

    if numbers.is_empty() {
        return Err(CellValue::Error(format!(
            "{func_name} requires at least one numeric value in the reference"
        )));
    }

    Ok(numbers)
}

pub(super) fn to_positive_index(
    value: f64,
    func_name: &str,
    arg_name: &str,
) -> StdResult<usize, CellValue> {
    if !value.is_finite() {
        return Err(CellValue::Error(format!(
            "{func_name} {arg_name} must be a finite positive integer"
        )));
    }
    let rounded = value.round();
    if (value - rounded).abs() > EPS {
        return Err(CellValue::Error(format!(
            "{func_name} {arg_name} must be an integer"
        )));
    }
    if rounded < 1.0 {
        return Err(CellValue::Error(format!(
            "{func_name} {arg_name} must be greater than or equal to 1"
        )));
    }
    Ok(rounded as usize)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::engine::test_helpers::TestEngine;

    fn num_expr(n: f64) -> Expr {
        if n == n.floor() {
            Expr::Literal(CellValue::Int(n as i64))
        } else {
            Expr::Literal(CellValue::Float(n))
        }
    }

    // ===== number_arg tests =====

    #[tokio::test]
    async fn test_number_arg_with_number() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let expr = num_expr(42.0);
        let result = number_arg(ctx, "Sheet1", &expr).await.unwrap();
        assert_eq!(result, Some(42.0));
    }

    #[tokio::test]
    async fn test_number_arg_with_string() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let expr = Expr::Literal(CellValue::String("abc".to_string()));
        let result = number_arg(ctx, "Sheet1", &expr).await.unwrap();
        assert_eq!(result, None);
    }

    // ===== collect_numeric_values tests =====

    #[test]
    fn test_collect_numeric_values_basic() {
        let values = vec![CellValue::Int(3), CellValue::Int(1), CellValue::Int(2)];
        let result = collect_numeric_values(values, "MEDIAN").unwrap();
        assert_eq!(result, vec![1.0, 2.0, 3.0]); // sorted
    }

    #[test]
    fn test_collect_numeric_values_with_floats() {
        let values = vec![
            CellValue::Float(3.5),
            CellValue::Int(1),
            CellValue::Float(2.5),
        ];
        let result = collect_numeric_values(values, "AVERAGE").unwrap();
        assert_eq!(result, vec![1.0, 2.5, 3.5]);
    }

    #[test]
    fn test_collect_numeric_values_ignores_non_numeric() {
        let values = vec![
            CellValue::Int(3),
            CellValue::String("abc".to_string()),
            CellValue::Int(1),
        ];
        let result = collect_numeric_values(values, "COUNT").unwrap();
        assert_eq!(result, vec![1.0, 3.0]);
    }

    #[test]
    fn test_collect_numeric_values_empty_result() {
        let values = vec![CellValue::String("abc".to_string()), CellValue::Bool(true)];
        let result = collect_numeric_values(values, "SUM");
        assert!(result.is_err());
        match result.unwrap_err() {
            CellValue::Error(e) => assert!(e.contains("requires at least one numeric value")),
            _ => panic!("Expected Error"),
        }
    }

    #[test]
    fn test_collect_numeric_values_with_error() {
        let values = vec![
            CellValue::Int(1),
            CellValue::Error("Some error".to_string()),
            CellValue::Int(2),
        ];
        let result = collect_numeric_values(values, "AVERAGE");
        assert!(result.is_err());
        match result.unwrap_err() {
            CellValue::Error(e) => assert_eq!(e, "Some error"),
            _ => panic!("Expected Error"),
        }
    }

    #[test]
    fn test_collect_numeric_values_with_nan() {
        let values = vec![
            CellValue::Int(1),
            CellValue::Float(f64::NAN),
            CellValue::Int(2),
        ];
        let result = collect_numeric_values(values, "AVERAGE");
        assert!(result.is_err());
        match result.unwrap_err() {
            CellValue::Error(e) => assert!(e.contains("invalid numeric value")),
            _ => panic!("Expected Error"),
        }
    }

    // ===== collect_numeric_values_unsorted tests =====

    #[test]
    fn test_collect_numeric_values_unsorted() {
        let values = vec![CellValue::Int(3), CellValue::Int(1), CellValue::Int(2)];
        let result = collect_numeric_values_unsorted(values, "MEDIAN").unwrap();
        assert_eq!(result, vec![3.0, 1.0, 2.0]); // unsorted
    }

    #[test]
    fn test_collect_numeric_values_unsorted_empty() {
        let values: Vec<CellValue> = vec![];
        let result = collect_numeric_values_unsorted(values, "SUM");
        assert!(result.is_err());
        match result.unwrap_err() {
            CellValue::Error(e) => assert!(e.contains("requires at least one numeric value")),
            _ => panic!("Expected Error"),
        }
    }

    // ===== to_positive_index tests =====

    #[test]
    fn test_to_positive_index_valid() {
        assert_eq!(to_positive_index(1.0, "LARGE", "k").unwrap(), 1);
        assert_eq!(to_positive_index(5.0, "LARGE", "k").unwrap(), 5);
        // Values within EPS of an integer are accepted
        assert_eq!(to_positive_index(5.0000000000001, "LARGE", "k").unwrap(), 5);
    }

    #[test]
    fn test_to_positive_index_zero() {
        let result = to_positive_index(0.0, "LARGE", "k");
        assert!(result.is_err());
        match result.unwrap_err() {
            CellValue::Error(e) => assert!(e.contains("must be greater than or equal to 1")),
            _ => panic!("Expected Error"),
        }
    }

    #[test]
    fn test_to_positive_index_negative() {
        let result = to_positive_index(-1.0, "LARGE", "k");
        assert!(result.is_err());
        match result.unwrap_err() {
            CellValue::Error(e) => assert!(e.contains("must be greater than or equal to 1")),
            _ => panic!("Expected Error"),
        }
    }

    #[test]
    fn test_to_positive_index_non_integer() {
        let result = to_positive_index(1.5, "LARGE", "k");
        assert!(result.is_err());
        match result.unwrap_err() {
            CellValue::Error(e) => assert!(e.contains("must be an integer")),
            _ => panic!("Expected Error"),
        }
    }

    #[test]
    fn test_to_positive_index_infinite() {
        let result = to_positive_index(f64::INFINITY, "LARGE", "k");
        assert!(result.is_err());
        match result.unwrap_err() {
            CellValue::Error(e) => assert!(e.contains("must be a finite positive integer")),
            _ => panic!("Expected Error"),
        }
    }

    #[test]
    fn test_to_positive_index_nan() {
        let result = to_positive_index(f64::NAN, "LARGE", "k");
        assert!(result.is_err());
        match result.unwrap_err() {
            CellValue::Error(e) => assert!(e.contains("must be a finite positive integer")),
            _ => panic!("Expected Error"),
        }
    }

    // ===== EPS constant =====

    #[test]
    fn test_eps_value() {
        assert_eq!(EPS, 1e-12);
    }
}
