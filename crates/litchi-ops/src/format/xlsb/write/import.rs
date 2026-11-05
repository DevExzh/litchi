//! XLSB CSV/TSV import: every field staged in one native cell-values edit of
//! the target worksheet and published by one commit (record 0266).

use litchi_core::FileFormat;
use litchi_core::sheet::WorkbookTrait as _;

use super::super::node::{Grid, Item};
use super::super::owner::SheetHandle;
use super::super::{Nav, Xlsb};
use super::{
    Commit, NativeCellOptions, NativeFormat, NativeScalar, NativeSelector, Reference,
    caller_value_error, cell_patch_container_bytes, cell_snapshot_budget, cell_transfer_limits,
    native_error, package_from_editor,
};
use crate::error::{Code, Error, Result};
use crate::import::{Origin, Summary, Table, Value, Written, serial};
use crate::path::PathArg;
use crate::transaction::SemanticLimits;

/// The XLSB grid: 1,048,576 rows by 16,384 columns.
const MAX_ROWS: u32 = 1 << 20;
const MAX_COLUMNS: u32 = 1 << 14;

/// The zero-based worksheet position and origin cell of an import target:
/// a worksheet (origin `A1`) or one cell.
pub(crate) fn origin(file: &Xlsb, path: &PathArg<'_>) -> Result<Origin> {
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    match node.item() {
        Item::Sheet(sheet) => Ok(Origin {
            sheet: sheet.position,
            row: 0,
            column: 0,
            cell: false,
        }),
        Item::Cell(_) => {
            let cell = node.cell_ref().ok_or_else(not_a_target)?;
            Ok(Origin {
                sheet: node.sheet_position(),
                row: cell.row(),
                column: cell.column(),
                cell: true,
            })
        },
        _ => Err(not_a_target()),
    }
}

/// Stage every field of `table` in one native edit and commit once.
pub(crate) fn import(
    file: &Xlsb,
    path: &PathArg<'_>,
    table: &Table,
    limits: &SemanticLimits,
) -> Result<(Commit, Summary)> {
    file.check_source()?;
    let result = (|| {
        let origin = origin(file, path)?;
        let Origin {
            sheet, row, column, ..
        } = origin;
        table
            .check_grid(row, column, MAX_ROWS, MAX_COLUMNS)
            .map_err(|error| error.with_format(FileFormat::Xlsb))?;
        let date1904 = file.is_1904_date_system();
        // Empty fields clear only stored cells, so the read grid is pinned
        // when the first one is met.
        let mut existing: Option<Grid> = None;
        let native_limits =
            cell_transfer_limits(cell_snapshot_budget(limits.max_snapshot_bytes())?, limits)?;
        let package = package_from_editor(file.editor_for_write()?);
        let mut edit = package
            .edit_cell_values_with_options(
                NativeSelector::from(sheet),
                NativeCellOptions::new().transfer_limits(native_limits),
            )
            .map_err(native_error)?
            .resolve()
            .map_err(native_error)?
            .ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "the XLSB worksheet position is out of range",
                )
                .with_format(FileFormat::Xlsb)
            })?;
        let mut summary = Summary::at(origin, table);
        for field in table.fields() {
            let (at_row, at_column) = (row + field.row, column + field.column);
            let reference = Reference::new(at_row, at_column).map_err(native_error)?;
            let written = match &field.value {
                Value::Empty => {
                    let grid = match &mut existing {
                        Some(grid) => grid,
                        slot @ None => {
                            let handle = SheetHandle::at(file, sheet)?.ok_or_else(|| {
                                Error::new(
                                    Code::OutOfRange,
                                    "the XLSB worksheet position is out of range",
                                )
                                .with_format(FileFormat::Xlsb)
                            })?;
                            slot.insert(Grid::from_handle(&handle)?)
                        },
                    };
                    if grid.record((at_row, at_column))?.is_some() {
                        edit.set_scalar(reference, NativeScalar::Blank)
                            .map_err(native_error)?;
                    }
                    Written::Empty
                },
                Value::Text(text) => {
                    edit.set_scalar(reference, NativeScalar::Text(table.text(text).to_owned()))
                        .map_err(|error| caller_value_error(error, "value"))?;
                    Written::Text
                },
                Value::Number { value, .. } => {
                    edit.set_scalar(reference, NativeScalar::Number(*value))
                        .map_err(|error| caller_value_error(error, "value"))?;
                    Written::Number
                },
                Value::Bool(value) => {
                    edit.set_scalar(reference, NativeScalar::Boolean(*value))
                        .map_err(native_error)?;
                    Written::Boolean
                },
                Value::Date {
                    days,
                    seconds,
                    time,
                    text,
                } => match serial(*days, *seconds, date1904) {
                    Some(serial) => {
                        edit.set_scalar(reference, NativeScalar::Number(serial))
                            .map_err(native_error)?;
                        let code = if *time {
                            "yyyy-mm-dd hh:mm:ss"
                        } else {
                            "yyyy-mm-dd"
                        };
                        let mut format = NativeFormat::default();
                        format.number_format = Some(code.to_owned());
                        edit.set_format(reference, &format).map_err(native_error)?;
                        Written::Date
                    },
                    None => {
                        edit.set_scalar(reference, NativeScalar::Text(table.text(text).to_owned()))
                            .map_err(native_error)?;
                        Written::DateAsText
                    },
                },
                Value::Formula(text) => {
                    edit.set_formula_text(reference, table.text(text))
                        .map_err(|error| caller_value_error(error, "formula"))?;
                    Written::Formula
                },
            };
            summary.count(written);
        }
        let native = edit.commit().map_err(native_error)?;
        let (next, patch) = native.into_parts();
        let retained_snapshot = cell_patch_container_bytes()
            .checked_add(patch.retained_bytes())
            .ok_or_else(|| super::snapshot_limit(usize::MAX, limits.max_snapshot_bytes()))?;
        let commit = Commit::from_cell_parts(
            file,
            next,
            vec![patch],
            (sheet, native_limits),
            retained_snapshot,
            limits.max_snapshot_bytes(),
            table.fields().len(),
        )?;
        Ok((commit, summary))
    })();
    let freshness = file.check_source();
    freshness.and(result)
}

fn not_a_target() -> Error {
    Error::new(
        Code::InvalidArgument,
        "import targets a worksheet or one cell of a worksheet",
    )
    .with_format(FileFormat::Xlsb)
    .with_suggestion("use a path such as /Sheet1 or /Sheet1/B2")
}
