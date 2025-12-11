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
