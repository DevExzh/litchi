use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, for_each_value_in_expr, is_blank};

pub(crate) async fn eval_count(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut count = 0u64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if matches!(
                v,
                CellValue::Int(_) | CellValue::Float(_) | CellValue::DateTime(_)
            ) {
                count += 1;
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Int(count as i64))
}

pub(crate) async fn eval_countblank(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut count = 0u64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if is_blank(v) {
                count += 1;
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Int(count as i64))
}

pub(crate) async fn eval_counta(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut count = 0u64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if !matches!(v, CellValue::Empty) {
                count += 1;
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Int(count as i64))
}
