//! XLS CSV/TSV import: the fields are staged in native `Package::edit`s of
//! at most 1,024 cells, each committed on the last, inside one unified
//! transaction (record 0266).
//!
//! BIFF8 has no unified number-format write yet, so inferred dates are kept
//! as their ISO text ([`Summary::dates_as_text`]). Existing formula cells are
//! refused, as the unified XLS cell setter refuses them.

use litchi_core::{FileFormat, Position};
use litchi_xls::cell_values::{Reference, Selector, Storage, Value as NativeValue};
use std::mem::size_of;

use super::super::node::Item;
use super::super::{Nav, Xls};
use super::{WriteCommit, finish_commit, formula_error, native_edit_error, native_value_error};
use crate::error::{Code, Error, Result};
use crate::import::{Origin, Summary, Table, Value, Written};
use crate::path::PathArg;
use crate::transaction::SemanticLimits;

/// The BIFF8 grid: 65,536 rows by 256 columns.
const MAX_ROWS: u32 = 1 << 16;
const MAX_COLUMNS: u32 = 1 << 8;

/// The zero-based worksheet position and origin cell of an import target:
/// a worksheet (origin `A1`) or one cell.
pub(crate) fn origin(file: &Xls, path: &PathArg<'_>) -> Result<Origin> {
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    match node.item() {
        Item::Sheet(sheet) => Ok(Origin {
            sheet: sheet.position,
            row: 0,
            column: 0,
            cell: false,
        }),
        Item::Cell(_) => {
            let (column, row) = node.cell().ok_or_else(not_a_target)?;
            Ok(Origin {
                sheet: node.sheet_position(),
                row,
                column,
                cell: true,
            })
        },
        _ => Err(not_a_target()),
    }
}

/// The cells one native BIFF8 edit stages before it commits: the leaf
/// admits 4,096 staged changes, and a text cell can stage a shared-string
/// resource beside its record.
const CHUNK: usize = 1_024;

/// Stage the fields of `table` in native edits of [`CHUNK`] cells, each
/// committed on the previous one's result.
pub(crate) fn import(
    file: &Xls,
    path: &PathArg<'_>,
    table: &Table,
    limits: &SemanticLimits,
) -> Result<(Vec<WriteCommit>, Summary)> {
    file.admit_edit_for_write()?;
    let origin = origin(file, path)?;
    let Origin {
        sheet, row, column, ..
    } = origin;
    table
        .check_grid(row, column, MAX_ROWS, MAX_COLUMNS)
        .map_err(|error| error.with_format(FileFormat::Xls))?;
    let selector = || Selector::Position(Position::new(sheet));
    let mut summary = Summary::at(origin, table);
    let mut commits: Vec<WriteCommit> = Vec::new();
    for chunk in table.fields().chunks(CHUNK) {
        let base = commits
            .last()
            .map_or_else(|| file.clone(), |commit| commit.after.clone());
        let used = commits.iter().fold(0usize, |total, commit| {
            total.saturating_add(commit.snapshot_bytes)
        });
        let remaining = limits.max_snapshot_bytes().saturating_sub(used);
        let chunk_limits = limits.with_snapshot_budget(remaining);
        let additional = chunk.iter().fold(0usize, |total, field| {
            let payload = match &field.value {
                Value::Text(text) | Value::Formula(text) => table.text(text).len(),
                Value::Date { text, .. } => table.text(text).len(),
                Value::Empty | Value::Number { .. } | Value::Bool(_) => 0,
            };
            total
                .saturating_add(payload)
                .saturating_add(size_of::<NativeValue>())
        });
        base.preflight_snapshot_bytes(remaining, additional)?;
        let mut edit = base
            .editable_package_for_write_with_snapshot_limit(remaining)?
            .edit()
            .map_err(native_edit_error)?;
        for field in chunk {
            let reference = Reference::new(row + field.row, column + field.column)
                .map_err(native_value_error)?;
            let stored = edit
                .source_cell(selector(), reference)
                .map_err(native_edit_error)?
                .map(|cell| cell.storage() == Storage::Formula);
            let (value, written) = match &field.value {
                Value::Empty => {
                    if stored.is_some() {
                        refuse_formula(stored)?;
                        edit.set_value(selector(), reference, NativeValue::Blank)
                            .map_err(native_edit_error)?;
                    }
                    summary.count(Written::Empty);
                    continue;
                },
                Value::Formula(text) => {
                    let text = table.text(text);
                    refuse_formula(stored)?;
                    if stored.is_some() {
                        edit.set_formula(selector(), reference, text)
                    } else {
                        edit.insert_formula(selector(), reference, text)
                    }
                    .map_err(formula_error)?;
                    summary.count(Written::Formula);
                    continue;
                },
                Value::Text(text) => (
                    NativeValue::Text(text_value(table.text(text))?),
                    Written::Text,
                ),
                Value::Number { value, .. } => (NativeValue::Number(*value), Written::Number),
                Value::Bool(value) => (NativeValue::Boolean(*value), Written::Boolean),
                Value::Date { text, .. } => (
                    NativeValue::Text(text_value(table.text(text))?),
                    Written::DateAsText,
                ),
            };
            match stored {
                Some(formula) => {
                    refuse_formula(Some(formula))?;
                    edit.set_value(selector(), reference, value)
                },
                None => edit.insert_cell(selector(), reference, value),
            }
            .map_err(native_edit_error)?;
            summary.count(written);
        }
        let native = edit.commit().map_err(native_edit_error)?;
        commits.push(finish_commit(&base, native, 1, &chunk_limits)?);
    }
    file.check_source()?;
    Ok((commits, summary))
}

/// BIFF8 text is at most 65,535 UTF-16 units; the parser already bounds a
/// field at 32,767 characters and refuses characters no cell can store.
fn text_value(text: &str) -> Result<String> {
    if text.encode_utf16().count() > 65_535 {
        return Err(
            Error::new(Code::InvalidValue, "the XLS cell text is too long")
                .with_format(FileFormat::Xls)
                .with_property("value"),
        );
    }
    Ok(text.to_owned())
}

fn refuse_formula(stored: Option<bool>) -> Result<()> {
    if stored == Some(true) {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "importing over an XLS formula cell requires formula-cache or formula-token editing",
        )
        .with_format(FileFormat::Xls)
        .with_property("value"));
    }
    Ok(())
}

fn not_a_target() -> Error {
    Error::new(
        Code::InvalidArgument,
        "import targets a worksheet or one cell of a worksheet",
    )
    .with_format(FileFormat::Xls)
    .with_suggestion("use a path such as /Sheet1 or /Sheet1/B2")
}
