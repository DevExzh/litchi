use crate::sheet::CellValue;

use super::super::parser::BinaryOp;
use super::{to_number, to_text};

pub(crate) fn eval_binary_op(op: BinaryOp, left: CellValue, right: CellValue) -> CellValue {
    match op {
        BinaryOp::Add | BinaryOp::Sub | BinaryOp::Mul | BinaryOp::Div => {
            let ln = match to_number(&left) {
                Some(n) => n,
                None => {
                    return CellValue::Error("Left operand is not numeric".to_string());
                },
            };

            let rn = match to_number(&right) {
                Some(n) => n,
                None => {
                    return CellValue::Error("Right operand is not numeric".to_string());
                },
            };

            let result = match op {
                BinaryOp::Add => ln + rn,
                BinaryOp::Sub => ln - rn,
                BinaryOp::Mul => ln * rn,
                BinaryOp::Div => {
                    if rn == 0.0 {
                        return CellValue::Error("Division by zero".to_string());
                    }
                    ln / rn
                },
                _ => unreachable!(),
            };

            // For now, always return a Float result. We can refine this later to
            // preserve integer types when possible.
            CellValue::Float(result)
        },
        BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Gt | BinaryOp::Ge | BinaryOp::Lt | BinaryOp::Le => {
            eval_comparison(op, left, right)
        },
    }
}

fn eval_comparison(op: BinaryOp, left: CellValue, right: CellValue) -> CellValue {
    if let CellValue::Error(e) = &left {
        return CellValue::Error(e.clone());
    }
    if let CellValue::Error(e) = &right {
        return CellValue::Error(e.clone());
    }

    let ln = to_number(&left);
    let rn = to_number(&right);

    let result = if let (Some(ln), Some(rn)) = (ln, rn) {
        match op {
            BinaryOp::Eq => ln == rn,
            BinaryOp::Ne => ln != rn,
            BinaryOp::Gt => ln > rn,
            BinaryOp::Ge => ln >= rn,
            BinaryOp::Lt => ln < rn,
            BinaryOp::Le => ln <= rn,
            _ => unreachable!(),
        }
    } else {
        let ls = to_text(&left);
        let rs = to_text(&right);
        match op {
            BinaryOp::Eq => ls == rs,
            BinaryOp::Ne => ls != rs,
            BinaryOp::Gt => ls > rs,
            BinaryOp::Ge => ls >= rs,
            BinaryOp::Lt => ls < rs,
            BinaryOp::Le => ls <= rs,
            _ => unreachable!(),
        }
    };

    CellValue::Bool(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::CellValue;
    use crate::sheet::eval::parser::BinaryOp;

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
            CellValue::Error(e) => assert_eq!(e, "Division by zero"),
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
            CellValue::Error(e) => assert_eq!(e, "Left operand is not numeric"),
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
            CellValue::Error(e) => assert_eq!(e, "Right operand is not numeric"),
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
