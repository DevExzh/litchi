use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, evaluate_expression};

pub(crate) async fn eval_choose(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 {
        return Ok(CellValue::Error(
            "CHOOSE expects at least 2 arguments (index_num, value1, ...)".to_string(),
        ));
    }

    let index_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let index = match super::super::to_number(&index_val) {
        Some(n) if n >= 1.0 => n.trunc() as usize,
        _ => {
            return Ok(CellValue::Error(
                "CHOOSE index_num must be a positive number".to_string(),
            ));
        },
    };

    let choices = args.len() - 1;
    if index == 0 || index > choices {
        return Ok(CellValue::Error(
            "CHOOSE index_num out of range".to_string(),
        ));
    }

    let choice_expr = &args[index];
    evaluate_expression(ctx, current_sheet, choice_expr).await
}
