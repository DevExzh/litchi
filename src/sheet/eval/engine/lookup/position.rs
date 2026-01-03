use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, flatten_range_expr};
use super::helpers::{ReferenceLookup, first_cell_from_expr};

pub(crate) async fn eval_column(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() > 1 {
        return Ok(CellValue::Error(
            "COLUMN expects at most 1 argument".to_string(),
        ));
    }

    let col = if args.is_empty() {
        match ctx.current_position() {
            Some((_, _, col)) => col,
            None => {
                return Ok(CellValue::Error(
                    "COLUMN cannot determine the current position".to_string(),
                ));
            },
        }
    } else {
        match first_cell_from_expr(ctx, current_sheet, &args[0]).await? {
            ReferenceLookup::Point((_row, col)) => col,
            ReferenceLookup::NameError(msg) => {
                return Ok(CellValue::Error(msg));
            },
            ReferenceLookup::NotReference => {
                return Ok(CellValue::Error(
                    "COLUMN expects a cell or range reference".to_string(),
                ));
            },
        }
    };

    Ok(CellValue::Int(col as i64))
}

pub(crate) async fn eval_row(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() > 1 {
        return Ok(CellValue::Error(
            "ROW expects at most 1 argument".to_string(),
        ));
    }

    let row = if args.is_empty() {
        match ctx.current_position() {
            Some((_, row, _)) => row,
            None => {
                return Ok(CellValue::Error(
                    "ROW cannot determine the current position".to_string(),
                ));
            },
        }
    } else {
        match first_cell_from_expr(ctx, current_sheet, &args[0]).await? {
            ReferenceLookup::Point((row, _col)) => row,
            ReferenceLookup::NameError(msg) => {
                return Ok(CellValue::Error(msg));
            },
            ReferenceLookup::NotReference => {
                return Ok(CellValue::Error(
                    "ROW expects a cell or range reference".to_string(),
                ));
            },
        }
    };

    Ok(CellValue::Int(row as i64))
}

pub(crate) async fn eval_rows(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error(
            "ROWS expects exactly 1 argument (array)".to_string(),
        ));
    }

    let range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    Ok(CellValue::Int(range.rows as i64))
}

pub(crate) async fn eval_columns(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error(
            "COLUMNS expects exactly 1 argument (array)".to_string(),
        ));
    }

    let range = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    Ok(CellValue::Int(range.cols as i64))
}
