//! Tab-separated text projection of worksheet rows, moved verbatim from the
//! facade's `sheet::workbook::append_cell_text`.
//!
//! The workbook handles include this file as a private submodule each
//! (`#[path]`), because `format/mod.rs` is shared by the parallel S5 steps.
//! S6 can promote it to one `format` module.

use litchi_core::sheet::CellValue;

/// Append one row: cells separated by tabs, ended by a newline.
pub(super) fn append_row(out: &mut String, row: &[CellValue]) {
    for (index, cell) in row.iter().enumerate() {
        if index > 0 {
            out.push('\t');
        }
        append_cell_text(out, cell);
    }
    out.push('\n');
}

fn append_cell_text(out: &mut String, cell: &CellValue) {
    match cell {
        CellValue::Empty => {},
        CellValue::Bool(value) => out.push_str(if *value { "TRUE" } else { "FALSE" }),
        CellValue::Int(value) => out.push_str(&value.to_string()),
        CellValue::Float(value) => out.push_str(&value.to_string()),
        CellValue::String(value) => out.push_str(value),
        CellValue::DateTime(value) => out.push_str(&value.to_string()),
        CellValue::Error(value) => out.push_str(value),
        CellValue::Formula {
            formula,
            cached_value,
            ..
        } => match cached_value.as_deref() {
            Some(CellValue::String(value)) => out.push_str(value),
            Some(CellValue::Int(value)) => out.push_str(&value.to_string()),
            Some(CellValue::Float(value)) => out.push_str(&value.to_string()),
            Some(CellValue::Bool(value)) => out.push_str(if *value { "TRUE" } else { "FALSE" }),
            _ => {
                out.push('=');
                out.push_str(formula);
            },
        },
    }
}
