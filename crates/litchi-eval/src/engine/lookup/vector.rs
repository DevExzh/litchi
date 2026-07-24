use std::cmp::Ordering;

use crate::parser::Expr;
use litchi_core::sheet::{CellValue, Result};

use super::super::{EvalCtx, evaluate_expression, flatten_range_expr, to_number};
use super::helpers::is_1d;

/// Sort-rank buckets used by Excel's approximate-match ordering:
/// empty < number < text < boolean. Error values are never compared.
fn order_rank(value: &CellValue) -> u8 {
    match value {
        CellValue::Empty => 0,
        CellValue::Int(_) | CellValue::Float(_) | CellValue::DateTime(_) => 1,
        CellValue::String(_) => 2,
        CellValue::Bool(_) => 3,
        CellValue::Error(_) | CellValue::Formula { .. } => 4,
    }
}

/// Compares two cell values the way Excel orders values for an ascending
/// approximate match (LOOKUP): numbers numerically, text case-insensitively,
/// FALSE before TRUE, and each kind ordered before the next.
fn lookup_cmp(a: &CellValue, b: &CellValue) -> Ordering {
    let (ra, rb) = (order_rank(a), order_rank(b));
    if ra != rb {
        return ra.cmp(&rb);
    }
    match (a, b) {
        (
            CellValue::Int(_) | CellValue::Float(_) | CellValue::DateTime(_),
            CellValue::Int(_) | CellValue::Float(_) | CellValue::DateTime(_),
        ) => match (to_number(a), to_number(b)) {
            (Some(x), Some(y)) => x.partial_cmp(&y).unwrap_or(Ordering::Equal),
            _ => Ordering::Equal,
        },
        (CellValue::String(s1), CellValue::String(s2)) => s1.to_uppercase().cmp(&s2.to_uppercase()),
        (CellValue::Bool(b1), CellValue::Bool(b2)) => b1.cmp(b2),
        _ => Ordering::Equal,
    }
}

/// LOOKUP(lookup_value, lookup_vector, [result_vector]) — vector form only.
///
/// The array form `LOOKUP(lookup_value, array)` is not implemented: it picks
/// its result from the last row/column of a two-dimensional array, and this
/// engine's `CellValue` has no array variant to pass such a value in.
pub(crate) async fn eval_lookup(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "LOOKUP expects 2 or 3 arguments (lookup_value, lookup_vector, [result_vector])"
                .to_string(),
        ));
    }

    let lookup_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = lookup_val {
        return Ok(lookup_val);
    }

    let lookup_vector = flatten_range_expr(ctx, current_sheet, &args[1]).await?;
    if !is_1d(&lookup_vector) {
        return Ok(CellValue::Error("#N/A".to_string()));
    }

    let result_values = if args.len() == 3 {
        let result_vector = flatten_range_expr(ctx, current_sheet, &args[2]).await?;
        if !is_1d(&result_vector) || result_vector.values.len() != lookup_vector.values.len() {
            return Ok(CellValue::Error("#N/A".to_string()));
        }
        result_vector.values
    } else {
        lookup_vector.values.clone()
    };

    // Approximate match: the vector is assumed ascending; the result is the
    // position of the largest value <= lookup_value. A linear scan keeps this
    // deterministic even when the assumption is violated (Excel's binary
    // search result is unspecified in that case). Error cells never match.
    let mut candidate: Option<usize> = None;
    for (idx, v) in lookup_vector.values.iter().enumerate() {
        if matches!(v, CellValue::Error(_)) {
            continue;
        }
        if lookup_cmp(v, &lookup_val) != Ordering::Greater {
            candidate = Some(idx);
        }
    }

    match candidate {
        Some(idx) => Ok(result_values[idx].clone()),
        // lookup_value is smaller than the first (smallest) vector value.
        None => Ok(CellValue::Error("#N/A".to_string())),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::engine::test_helpers::TestEngine;
    use crate::parser::ast::RangeRef;

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

    fn range_expr(sheet: &str, start_row: u32, start_col: u32, end_row: u32, end_col: u32) -> Expr {
        Expr::Range(RangeRef {
            sheet: sheet.to_string(),
            start_row,
            start_col,
            end_row,
            end_col,
        })
    }

    fn ints(values: &[i64]) -> Vec<CellValue> {
        values.iter().map(|v| CellValue::Int(*v)).collect()
    }

    #[tokio::test]
    async fn test_lookup_two_arg_form() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Ascending vector; LOOKUP(3) finds 3 itself.
        engine.add_range("Sheet1", 1, 1, 5, 1, ints(&[1, 2, 3, 4, 5]));
        let args = vec![num_expr(3.0), range_expr("Sheet1", 1, 1, 5, 1)];
        let result = eval_lookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(3));
    }

    #[tokio::test]
    async fn test_lookup_approximate_match() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // LOOKUP(4.5) returns the largest value <= 4.5, i.e. 4.
        engine.add_range("Sheet1", 1, 1, 5, 1, ints(&[1, 2, 3, 4, 5]));
        let args = vec![num_expr(4.5), range_expr("Sheet1", 1, 1, 5, 1)];
        let result = eval_lookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(4));
    }

    #[tokio::test]
    async fn test_lookup_greater_than_all() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.add_range("Sheet1", 1, 1, 3, 1, ints(&[10, 20, 30]));
        let args = vec![num_expr(999.0), range_expr("Sheet1", 1, 1, 3, 1)];
        let result = eval_lookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(30));
    }

    #[tokio::test]
    async fn test_lookup_below_first_returns_na() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.add_range("Sheet1", 1, 1, 3, 1, ints(&[10, 20, 30]));
        let args = vec![num_expr(5.0), range_expr("Sheet1", 1, 1, 3, 1)];
        let result = eval_lookup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#N/A"),
            _ => panic!("Expected #N/A, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_lookup_with_result_vector() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.add_range("Sheet1", 1, 1, 3, 1, ints(&[1, 2, 3]));
        engine.add_range(
            "Sheet1",
            1,
            2,
            3,
            1,
            vec![
                CellValue::String("one".to_string()),
                CellValue::String("two".to_string()),
                CellValue::String("three".to_string()),
            ],
        );
        // LOOKUP(2.5, A1:A3, B1:B3) -> "two"
        let args = vec![
            num_expr(2.5),
            range_expr("Sheet1", 1, 1, 3, 1),
            range_expr("Sheet1", 1, 2, 3, 2),
        ];
        let result = eval_lookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("two".to_string()));
    }

    #[tokio::test]
    async fn test_lookup_result_vector_size_mismatch() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.add_range("Sheet1", 1, 1, 3, 1, ints(&[1, 2, 3]));
        engine.add_range("Sheet1", 1, 2, 2, 1, ints(&[10, 20]));
        let args = vec![
            num_expr(2.0),
            range_expr("Sheet1", 1, 1, 3, 1),
            range_expr("Sheet1", 1, 2, 2, 2),
        ];
        let result = eval_lookup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#N/A"),
            _ => panic!("Expected #N/A, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_lookup_non_vector_range() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.add_range("Sheet1", 1, 1, 2, 2, ints(&[1, 2, 3, 4]));
        let args = vec![num_expr(2.0), range_expr("Sheet1", 1, 1, 2, 2)];
        let result = eval_lookup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#N/A"),
            _ => panic!("Expected #N/A, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_lookup_text_values_case_insensitive() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.add_range(
            "Sheet1",
            1,
            1,
            3,
            1,
            vec![
                CellValue::String("apple".to_string()),
                CellValue::String("banana".to_string()),
                CellValue::String("cherry".to_string()),
            ],
        );
        // Case-insensitive: "BANANA" matches "banana".
        let args = vec![str_expr("BANANA"), range_expr("Sheet1", 1, 1, 3, 1)];
        let result = eval_lookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("banana".to_string()));
    }

    #[tokio::test]
    async fn test_lookup_text_sorts_after_numbers() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Mixed vector: numbers sort before text, so LOOKUP(99) skips the text.
        engine.add_range(
            "Sheet1",
            1,
            1,
            4,
            1,
            vec![
                CellValue::Int(1),
                CellValue::Int(2),
                CellValue::Int(3),
                CellValue::String("zebra".to_string()),
            ],
        );
        let args = vec![num_expr(99.0), range_expr("Sheet1", 1, 1, 4, 1)];
        let result = eval_lookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(3));
    }

    #[tokio::test]
    async fn test_lookup_horizontal_vector() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.add_range("Sheet1", 1, 1, 1, 3, ints(&[5, 10, 15]));
        let args = vec![num_expr(12.0), range_expr("Sheet1", 1, 1, 1, 3)];
        let result = eval_lookup(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(10));
    }

    #[tokio::test]
    async fn test_lookup_error_lookup_value_propagates() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.add_range("Sheet1", 1, 1, 2, 1, ints(&[1, 2]));
        let args = vec![
            Expr::Literal(CellValue::Error("#DIV/0!".to_string())),
            range_expr("Sheet1", 1, 1, 2, 1),
        ];
        let result = eval_lookup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#DIV/0!"),
            _ => panic!("Expected #DIV/0!, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_lookup_wrong_arg_count() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_lookup(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 or 3")),
            _ => panic!("Expected Error"),
        }
    }
}
