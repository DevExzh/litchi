use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, flatten_range_expr, to_number};

pub(crate) async fn eval_index(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "INDEX expects 2 or 3 arguments (array, row_num, [column_num])".to_string(),
        ));
    }

    let array = flatten_range_expr(ctx, current_sheet, &args[0]).await?;

    let row_val = super::super::evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let row_num = match to_number(&row_val) {
        Some(n) if n >= 1.0 => n as i64,
        _ => {
            return Ok(CellValue::Error(
                "INDEX row_num must be a positive number".to_string(),
            ));
        },
    };

    let col_num = if args.len() == 3 {
        let col_val = super::super::evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match to_number(&col_val) {
            Some(n) if n >= 1.0 => n as i64,
            _ => {
                return Ok(CellValue::Error(
                    "INDEX column_num must be a positive number".to_string(),
                ));
            },
        }
    } else {
        1
    };

    let rows = array.rows as i64;
    let cols = array.cols as i64;

    if row_num < 1 || row_num > rows || col_num < 1 || col_num > cols {
        return Ok(CellValue::Error(
            "INDEX row_num/column_num out of bounds for array".to_string(),
        ));
    }

    let idx = ((row_num - 1) * cols + (col_num - 1)) as usize;
    Ok(array.values[idx].clone())
}
