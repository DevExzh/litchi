use std::cmp::min;

use crate::sheet::eval::parser::{Expr, RangeRef};
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, FlatRange, ResolvedName, to_number, to_text};

pub(super) enum ReferenceLookup {
    Point((u32, u32)),
    NameError(String),
    NotReference,
}

pub(super) async fn first_cell_from_expr(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    expr: &Expr,
) -> Result<ReferenceLookup> {
    let lookup = match expr {
        Expr::Reference { row, col, .. } => ReferenceLookup::Point((*row, *col)),
        Expr::Range(range) => ReferenceLookup::Point(range_first_cell(range)),
        Expr::Name(name) => match ctx.resolve_name(current_sheet, name.as_str())? {
            Some(ResolvedName::Cell { row, col, .. }) => ReferenceLookup::Point((row, col)),
            Some(ResolvedName::Range(range)) => ReferenceLookup::Point(range_first_cell(&range)),
            None => ReferenceLookup::NameError(format!("Unknown name: {}", name)),
        },
        _ => ReferenceLookup::NotReference,
    };

    Ok(lookup)
}

fn range_first_cell(range: &RangeRef) -> (u32, u32) {
    let row = min(range.start_row, range.end_row);
    let col = min(range.start_col, range.end_col);
    (row, col)
}

pub(super) fn is_1d(range: &FlatRange) -> bool {
    range.rows == 1 || range.cols == 1
}

pub(super) fn find_exact_match_index(
    lookup_val: &CellValue,
    values: &[CellValue],
) -> Option<usize> {
    for (idx, v) in values.iter().enumerate() {
        if values_equal(lookup_val, v) {
            return Some(idx);
        }
    }
    None
}

pub(super) fn values_equal(a: &CellValue, b: &CellValue) -> bool {
    match (to_number(a), to_number(b)) {
        (Some(x), Some(y)) => x == y,
        _ => to_text(a) == to_text(b),
    }
}
