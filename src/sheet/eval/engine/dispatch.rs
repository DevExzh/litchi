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
