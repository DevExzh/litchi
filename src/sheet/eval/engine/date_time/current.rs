use crate::sheet::eval::engine::EvalCtx;
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use chrono::Utc;

use super::helpers::{date_to_excel_serial_1900, datetime_to_excel_serial_1900};

pub(crate) async fn eval_today(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("TODAY expects 0 arguments".to_string()));
    }

    let now = Utc::now().date_naive();
    let serial = date_to_excel_serial_1900(now);
    Ok(CellValue::DateTime(serial))
}

pub(crate) async fn eval_now(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("NOW expects 0 arguments".to_string()));
    }

    let now = Utc::now().naive_utc();
    let serial = datetime_to_excel_serial_1900(now);
    Ok(CellValue::DateTime(serial))
}
