use litchi_core::sheet::CellValue;

use super::super::parser::BinaryOp;
use super::{to_number, to_text};

pub(crate) fn eval_binary_op(op: BinaryOp, left: CellValue, right: CellValue) -> CellValue {
    // A spreadsheet error operand propagates, the left one first.
    if let CellValue::Error(error) = &left {
        return CellValue::Error(error.clone());
    }
    if let CellValue::Error(error) = &right {
        return CellValue::Error(error.clone());
    }
    match op {
        BinaryOp::Add | BinaryOp::Sub | BinaryOp::Mul | BinaryOp::Div => {
            let (Some(ln), Some(rn)) = (arithmetic(&left), arithmetic(&right)) else {
                return CellValue::Error("#VALUE!".to_string());
            };

            let result = match op {
                BinaryOp::Add => ln + rn,
                BinaryOp::Sub => ln - rn,
                BinaryOp::Mul => ln * rn,
                BinaryOp::Div => {
                    if rn == 0.0 {
                        return CellValue::Error("#DIV/0!".to_string());
                    }
                    ln / rn
                },
                _ => unreachable!(),
            };
            if !result.is_finite() {
                return CellValue::Error("#NUM!".to_string());
            }

            // For now, always return a Float result. We can refine this later to
            // preserve integer types when possible.
            CellValue::Float(result)
        },
        BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Gt | BinaryOp::Ge | BinaryOp::Lt | BinaryOp::Le => {
            eval_comparison(op, left, right)
        },
    }
}

/// An arithmetic operand as Excel coerces it: a blank is 0, a boolean 1 or
/// 0, and text only when it spells a number.
fn arithmetic(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Empty => Some(0.0),
        CellValue::Bool(value) => Some(if *value { 1.0 } else { 0.0 }),
        CellValue::String(text) => text.trim().parse::<f64>().ok().filter(|n| n.is_finite()),
        other => to_number(other),
    }
}

/// Excel's comparison: numbers sort before text, text before booleans; text
/// compares without case; a blank compares as 0, empty text or FALSE,
/// whichever matches the other operand.
fn eval_comparison(op: BinaryOp, left: CellValue, right: CellValue) -> CellValue {
    use std::cmp::Ordering;

    #[derive(PartialEq, PartialOrd)]
    enum Key {
        Number(f64),
        Text(String),
        Bool(bool),
    }
    fn rank(key: &Key) -> u8 {
        match key {
            Key::Number(_) => 0,
            Key::Text(_) => 1,
            Key::Bool(_) => 2,
        }
    }
    fn key(value: &CellValue, other: &CellValue) -> Key {
        match value {
            CellValue::Empty => match other {
                CellValue::String(_) => Key::Text(String::new()),
                CellValue::Bool(_) => Key::Bool(false),
                _ => Key::Number(0.0),
            },
            CellValue::Bool(value) => Key::Bool(*value),
            CellValue::String(text) => Key::Text(text.to_lowercase()),
            other => match to_number(other) {
                Some(number) => Key::Number(number),
                None => Key::Text(to_text(other).to_lowercase()),
            },
        }
    }

    let (lk, rk) = (key(&left, &right), key(&right, &left));
    let ordering = match (&lk, &rk) {
        (Key::Number(l), Key::Number(r)) => l.partial_cmp(r).unwrap_or(Ordering::Equal),
        (Key::Text(l), Key::Text(r)) => l.cmp(r),
        (Key::Bool(l), Key::Bool(r)) => l.cmp(r),
        _ => rank(&lk).cmp(&rank(&rk)),
    };
    let result = match op {
        BinaryOp::Eq => ordering == Ordering::Equal,
        BinaryOp::Ne => ordering != Ordering::Equal,
        BinaryOp::Gt => ordering == Ordering::Greater,
        BinaryOp::Ge => ordering != Ordering::Less,
        BinaryOp::Lt => ordering == Ordering::Less,
        BinaryOp::Le => ordering != Ordering::Greater,
        _ => unreachable!(),
    };

    CellValue::Bool(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::BinaryOp;
    use litchi_core::sheet::CellValue;

    #[test]
    fn excel_coercion_and_ordering() {
        let num = |value: CellValue| match value {
            CellValue::Float(value) => value,
            other => panic!("expected a number, got {other:?}"),
        };
        assert_eq!(
            num(eval_binary_op(
                BinaryOp::Add,
                CellValue::Empty,
                CellValue::Int(1)
            )),
            1.0
        );
        assert_eq!(
            num(eval_binary_op(
                BinaryOp::Mul,
                CellValue::Bool(true),
                CellValue::String(" 2.5 ".into())
            )),
            2.5
        );
        assert_eq!(
            eval_binary_op(
                BinaryOp::Add,
                CellValue::Error("#N/A".into()),
                CellValue::Int(1)
            ),
            CellValue::Error("#N/A".into())
        );
        let cmp = |op, l, r| eval_binary_op(op, l, r) == CellValue::Bool(true);
        assert!(cmp(
            BinaryOp::Eq,
            CellValue::String("abc".into()),
            CellValue::String("ABC".into())
        ));
        assert!(cmp(
            BinaryOp::Lt,
            CellValue::Int(99),
            CellValue::String("1".into())
        ));
        assert!(cmp(
            BinaryOp::Gt,
            CellValue::Bool(false),
            CellValue::String("TRUE".into())
        ));
        assert!(cmp(BinaryOp::Eq, CellValue::Empty, CellValue::Int(0)));
        assert!(cmp(BinaryOp::Eq, CellValue::Empty, CellValue::Bool(false)));
        assert!(cmp(
            BinaryOp::Eq,
            CellValue::Empty,
            CellValue::String(String::new())
        ));
    }

    #[test]
    fn test_eval_binary_op_add() {
        let result = eval_binary_op(BinaryOp::Add, CellValue::Int(5), CellValue::Int(3));
        match result {
            CellValue::Float(v) => assert_eq!(v, 8.0),
            _ => panic!("Expected Float result"),
        }
    }

    #[test]
    fn test_eval_binary_op_sub() {
        let result = eval_binary_op(BinaryOp::Sub, CellValue::Int(10), CellValue::Int(4));
        match result {
            CellValue::Float(v) => assert_eq!(v, 6.0),
            _ => panic!("Expected Float result"),
        }
    }

    #[test]
    fn test_eval_binary_op_mul() {
        let result = eval_binary_op(BinaryOp::Mul, CellValue::Int(6), CellValue::Int(7));
        match result {
            CellValue::Float(v) => assert_eq!(v, 42.0),
            _ => panic!("Expected Float result"),
        }
    }

    #[test]
    fn test_eval_binary_op_div() {
        let result = eval_binary_op(BinaryOp::Div, CellValue::Int(15), CellValue::Int(3));
        match result {
            CellValue::Float(v) => assert_eq!(v, 5.0),
            _ => panic!("Expected Float result"),
        }
    }

    #[test]
    fn test_eval_binary_op_div_by_zero() {
        let result = eval_binary_op(BinaryOp::Div, CellValue::Int(10), CellValue::Int(0));
        match result {
            CellValue::Error(e) => assert_eq!(e, "#DIV/0!"),
            _ => panic!("Expected Error result"),
        }
    }

    #[test]
    fn test_eval_binary_op_non_numeric_left() {
        let result = eval_binary_op(
            BinaryOp::Add,
            CellValue::String("abc".to_string()),
            CellValue::Int(5),
        );
        match result {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            _ => panic!("Expected Error result"),
        }
    }

    #[test]
    fn test_eval_binary_op_non_numeric_right() {
        let result = eval_binary_op(
            BinaryOp::Add,
            CellValue::Int(5),
            CellValue::String("abc".to_string()),
        );
        match result {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            _ => panic!("Expected Error result"),
        }
    }

    #[test]
    fn test_eval_comparison_eq_numbers() {
        let result = eval_binary_op(BinaryOp::Eq, CellValue::Int(5), CellValue::Int(5));
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool result"),
        }
    }

    #[test]
    fn test_eval_comparison_ne_numbers() {
        let result = eval_binary_op(BinaryOp::Ne, CellValue::Int(5), CellValue::Int(3));
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool result"),
        }
    }

    #[test]
    fn test_eval_comparison_gt_numbers() {
        let result = eval_binary_op(BinaryOp::Gt, CellValue::Int(10), CellValue::Int(5));
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool result"),
        }
    }

    #[test]
    fn test_eval_comparison_lt_numbers() {
        let result = eval_binary_op(BinaryOp::Lt, CellValue::Int(3), CellValue::Int(7));
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool result"),
        }
    }

    #[test]
    fn test_eval_comparison_ge_numbers() {
        let result = eval_binary_op(BinaryOp::Ge, CellValue::Int(5), CellValue::Int(5));
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool result"),
        }
    }

    #[test]
    fn test_eval_comparison_le_numbers() {
        let result = eval_binary_op(BinaryOp::Le, CellValue::Int(3), CellValue::Int(5));
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool result"),
        }
    }

    #[test]
    fn test_eval_comparison_eq_strings() {
        let result = eval_binary_op(
            BinaryOp::Eq,
            CellValue::String("hello".to_string()),
            CellValue::String("hello".to_string()),
        );
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool result"),
        }
    }

    #[test]
    fn test_eval_comparison_gt_strings() {
        let result = eval_binary_op(
            BinaryOp::Gt,
            CellValue::String("zebra".to_string()),
            CellValue::String("apple".to_string()),
        );
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool result"),
        }
    }

    #[test]
    fn test_eval_comparison_mixed_types() {
        let result = eval_binary_op(BinaryOp::Eq, CellValue::Int(5), CellValue::Float(5.0));
        match result {
            CellValue::Bool(v) => assert!(v),
            _ => panic!("Expected Bool result"),
        }
    }

    #[test]
    fn test_eval_comparison_with_error_left() {
        let result = eval_binary_op(
            BinaryOp::Eq,
            CellValue::Error("#REF!".to_string()),
            CellValue::Int(5),
        );
        match result {
            CellValue::Error(e) => assert_eq!(e, "#REF!"),
            _ => panic!("Expected Error result"),
        }
    }

    #[test]
    fn test_eval_comparison_with_error_right() {
        let result = eval_binary_op(
            BinaryOp::Eq,
            CellValue::Int(5),
            CellValue::Error("#VALUE!".to_string()),
        );
        match result {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            _ => panic!("Expected Error result"),
        }
    }
}
