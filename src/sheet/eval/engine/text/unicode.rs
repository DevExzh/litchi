use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_text};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

pub(crate) async fn eval_unichar(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("UNICHAR expects 1 argument".to_string()));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let Some(n) = crate::sheet::eval::engine::to_number(&val) {
        let code = n.trunc() as u32;
        if code == 0 {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }
        match char::from_u32(code) {
            Some(c) => Ok(CellValue::String(c.to_string())),
            None => Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        Ok(CellValue::Error("#VALUE!".to_string()))
    }
}

pub(crate) async fn eval_unicode(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("UNICODE expects 1 argument".to_string()));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&val);
    if let Some(c) = text.chars().next() {
        Ok(CellValue::Int(c as u32 as i64))
    } else {
        Ok(CellValue::Error("#VALUE!".to_string()))
    }
}
