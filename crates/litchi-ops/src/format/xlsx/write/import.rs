//! XLSX CSV/TSV import: every field staged in one native workbook edit and
//! published by one commit (record 0266).

use litchi_core::FileFormat;
use litchi_core::sheet::WorkbookTrait as _;
use litchi_xlsx::{self as xlsx, Address};

use super::super::node::Item;
use super::super::{Nav, Xlsx};
use super::{Commit, native_error, validate_xml_text};
use crate::error::{Code, Error, Result};
use crate::import::{Origin, Summary, Table, Value, Written, serial};
use crate::path::PathArg;
use crate::transaction::SemanticLimits;

/// The XLSX grid: 1,048,576 rows by 16,384 columns.
const MAX_ROWS: u32 = 1 << 20;
const MAX_COLUMNS: u32 = 1 << 14;

/// The zero-based worksheet position and origin cell of an import target:
/// a worksheet (origin `A1`) or one cell.
pub(crate) fn origin(file: &Xlsx, path: &PathArg<'_>) -> Result<Origin> {
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    match node.item() {
        Item::Sheet(sheet) if sheet.handle.is_grid() => Ok(Origin {
            sheet: sheet.position,
            row: 0,
            column: 0,
            cell: false,
        }),
        Item::Cell(cell) => Ok(Origin {
            sheet: node.sheet_position(),
            row: cell.at.row().get(),
            column: cell.at.column().get(),
            cell: true,
        }),
        _ => Err(not_a_target()),
    }
}

/// Stage every field of `table` and commit once.
pub(crate) fn import(
    file: &Xlsx,
    path: &PathArg<'_>,
    table: &Table,
    _limits: &SemanticLimits,
) -> Result<(Commit, Summary)> {
    file.check_source()?;
    let result = (|| {
        let origin = origin(file, path)?;
        let Origin {
            sheet, row, column, ..
        } = origin;
        table
            .check_grid(row, column, MAX_ROWS, MAX_COLUMNS)
            .map_err(|error| error.with_format(FileFormat::Xlsx))?;
        let date1904 = file.is_1904_date_system();
        let source = file.editor_for_write_with_snapshot_limit(_limits.max_snapshot_bytes())?;
        let mut edit = super::edit_with_snapshot_limit(&source, _limits.max_snapshot_bytes())?;
        let mut summary = Summary::at(origin, table);
        {
            let mut worksheet = edit.sheet(sheet).map_err(native_error)?.ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "the XLSX worksheet position is out of range",
                )
                .with_format(FileFormat::Xlsx)
            })?;
            for field in table.fields() {
                let address =
                    Address::at(row + field.row, column + field.column).map_err(|_| {
                        Error::new(
                            Code::OutOfRange,
                            "the imported cell is outside the XLSX grid",
                        )
                        .with_format(FileFormat::Xlsx)
                    })?;
                let written = match &field.value {
                    Value::Empty => {
                        worksheet.clear(address).map_err(native_error)?;
                        Written::Empty
                    },
                    Value::Text(text) => {
                        let text = table.text(text);
                        validate_xml_text(text, "value")?;
                        worksheet
                            .set(address, xlsx::cell::Value::text(text))
                            .map_err(native_error)?;
                        Written::Text
                    },
                    Value::Number { lexical, .. } => {
                        let number =
                            xlsx::cell::Number::new(table.text(lexical)).map_err(native_error)?;
                        worksheet
                            .set(address, xlsx::cell::Value::from(number))
                            .map_err(native_error)?;
                        Written::Number
                    },
                    Value::Bool(value) => {
                        worksheet.set(address, *value).map_err(native_error)?;
                        Written::Boolean
                    },
                    Value::Date {
                        days,
                        seconds,
                        time,
                        text,
                    } => match serial(*days, *seconds, date1904) {
                        Some(serial) => {
                            let number =
                                xlsx::cell::Number::try_from(serial).map_err(native_error)?;
                            worksheet
                                .set(address, xlsx::cell::Value::from(number))
                                .map_err(native_error)?;
                            let code = if *time {
                                "yyyy-mm-dd hh:mm:ss"
                            } else {
                                "yyyy-mm-dd"
                            };
                            worksheet
                                .number_format(address, Some(code))
                                .map_err(native_error)?;
                            Written::Date
                        },
                        None => {
                            worksheet
                                .set(address, xlsx::cell::Value::text(table.text(text)))
                                .map_err(native_error)?;
                            Written::DateAsText
                        },
                    },
                    Value::Formula(text) => {
                        let text = table.text(text);
                        validate_xml_text(text, "formula")?;
                        let formula = xlsx::Formula::new(text).map_err(native_error)?;
                        worksheet.set(address, formula).map_err(native_error)?;
                        Written::Formula
                    },
                };
                summary.count(written);
            }
        }
        let native = edit.commit().map_err(native_error)?;
        Ok((
            Commit::from_native_with_limit(file, source, native, _limits.max_snapshot_bytes())?,
            summary,
        ))
    })();
    let freshness = file.check_source();
    freshness.and(result)
}

fn not_a_target() -> Error {
    Error::new(
        Code::InvalidArgument,
        "import targets a worksheet or one cell of a worksheet",
    )
    .with_format(FileFormat::Xlsx)
    .with_suggestion("use a path such as /Sheet1 or /Sheet1/B2")
}
