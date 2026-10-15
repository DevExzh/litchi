//! ODS CSV/TSV import (record 0295, record 0266 follow-up F8): every field
//! staged through the worksheet family of one leaf edit and published by
//! one commit.
//!
//! Inferred values take their native ODF types: numbers are `float`
//! cells, `TRUE`/`FALSE` `boolean` cells, and ISO dates `date` cells with
//! their `office:date-value` (an ODF date needs no serial or number
//! format). An existing cell keeps its style; an empty field clears a
//! stored cell's content. Formulas need translation into OpenFormula
//! syntax and are refused (`unsupported_capability`).

use std::collections::BTreeMap;

use litchi_core::FileFormat;
use litchi_ods::{Cell, CellValue};

use super::super::node::Item;
use super::super::{Nav, Ods};
use super::{Commit, finish, stage_cells};
use crate::error::{Code, Error, Result};
use crate::import::{Origin, Summary, Table, Value, Written};
use crate::path::PathArg;
use crate::transaction::SemanticLimits;

/// The ODF spreadsheet grid as LibreOffice bounds it.
const MAX_ROWS: u32 = 1 << 20;
const MAX_COLUMNS: u32 = 1 << 14;

/// The zero-based sheet position and origin cell of an import target: a
/// sheet (origin `A1`) or one cell.
pub(crate) fn origin(file: &Ods, path: &PathArg<'_>) -> Result<Origin> {
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    let sheet = node.sheet_position() as usize;
    match node.item() {
        Item::Sheet(_) => Ok(Origin {
            sheet,
            row: 0,
            column: 0,
            cell: false,
        }),
        Item::Cell(cell) => Ok(Origin {
            sheet,
            row: cell.at.0,
            column: cell.at.1,
            cell: true,
        }),
        _ => Err(not_a_target()),
    }
}

/// Stage every field of `table` in one leaf edit and commit once.
pub(crate) fn import(
    file: &Ods,
    path: &PathArg<'_>,
    table: &Table,
    limits: &SemanticLimits,
) -> Result<(Commit, Summary)> {
    file.check_source()?;
    let result = (|| {
        file.admit_edit()?;
        let origin = origin(file, path)?;
        let Origin {
            sheet, row, column, ..
        } = origin;
        table
            .check_grid(row, column, MAX_ROWS, MAX_COLUMNS)
            .map_err(|error| error.with_format(FileFormat::Ods))?;
        let sheets = file.worksheets()?;
        let data = sheets.get(sheet).ok_or_else(|| {
            Error::new(Code::OutOfRange, "the ODS sheet position is out of range")
        })?;
        let mut cells: BTreeMap<(u32, u32), Cell> = BTreeMap::new();
        let mut summary = Summary::at(origin, table);
        for field in table.fields() {
            let at = (row + field.row, column + field.column);
            let existing = data.cell(at.0 as usize, at.1 as usize);
            let (value, text, written) = match &field.value {
                Value::Empty => {
                    if existing.is_some_and(|cell| {
                        !matches!(cell.value, CellValue::Empty) || cell.formula.is_some()
                    }) {
                        let mut cell = fresh(existing);
                        cell.value = CellValue::Empty;
                        cell.text = String::new();
                        cells.insert(at, cell);
                    }
                    summary.count(Written::Empty);
                    continue;
                },
                Value::Text(range) => {
                    let text = table.text(range).to_owned();
                    (CellValue::Text(text.clone()), text, Written::Text)
                },
                Value::Number { value, lexical } => (
                    CellValue::Number(*value),
                    table.text(lexical).to_owned(),
                    Written::Number,
                ),
                Value::Bool(value) => (
                    CellValue::Boolean(*value),
                    if *value { "TRUE" } else { "FALSE" }.to_owned(),
                    Written::Boolean,
                ),
                Value::Date {
                    days,
                    seconds,
                    time,
                    ..
                } => {
                    let lexical = date_value(*days, *seconds, *time).ok_or_else(|| {
                        Error::new(
                            Code::InvalidValue,
                            "the imported date is outside the ODF range",
                        )
                    })?;
                    (CellValue::Date(lexical.clone()), lexical, Written::Date)
                },
                Value::Formula(_) => {
                    return Err(Error::new(
                        Code::UnsupportedCapability,
                        "ODS import writes values; formulas need OpenFormula translation",
                    )
                    .with_suggestion("import without `formulas(true)` (record 0295 follow-up)"));
                },
            };
            let mut cell = fresh(existing);
            cell.value = value;
            cell.text = text;
            cells.insert(at, cell);
            summary.count(written);
        }
        let mut edit = file.leaf_edit()?;
        let operations = cells.len();
        let mut staged = BTreeMap::new();
        if let Ok(position) = u32::try_from(sheet) {
            staged.insert(position, cells);
        }
        stage_cells(&mut edit, staged)?;
        Ok((finish(file, edit, operations, limits)?, summary))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Ods));
    file.check_source()?;
    result
}

/// A single cell to write over `existing`: its style and range source
/// kept, its links, formula and merge role dropped with the old content.
fn fresh(existing: Option<&Cell>) -> Cell {
    let mut cell = Cell::new(CellValue::Empty, String::new());
    if let Some(existing) = existing {
        cell.style_name.clone_from(&existing.style_name);
        cell.merge = existing.merge;
    }
    cell
}

/// The ODF `date` lexical value of a proleptic Gregorian date: days from
/// 1970-01-01 and seconds of the day.
fn date_value(days: i64, seconds: f64, time: bool) -> Option<String> {
    // Howard Hinnant's civil-from-days.
    let z = days.checked_add(719_468)?;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = if mp < 10 { mp + 3 } else { mp - 9 };
    let year = yoe + era * 400 + i64::from(month <= 2);
    if !(1..=9999).contains(&year) {
        return None;
    }
    let mut lexical = format!("{year:04}-{month:02}-{day:02}");
    if time {
        let whole = seconds.trunc();
        #[allow(
            clippy::cast_possible_truncation,
            clippy::cast_sign_loss,
            reason = "the seconds of a day are below 86,400"
        )]
        let whole = whole.clamp(0.0, 86_399.0) as u32;
        lexical.push_str(&format!(
            "T{:02}:{:02}:{:02}",
            whole / 3_600,
            whole / 60 % 60,
            whole % 60
        ));
        let fraction = seconds - seconds.trunc();
        if fraction > 0.0 {
            let digits = format!("{fraction:.9}");
            let digits = digits.trim_start_matches('0').trim_end_matches('0');
            if digits.len() > 1 {
                lexical.push_str(digits);
            }
        }
    }
    Some(lexical)
}

fn not_a_target() -> Error {
    Error::new(
        Code::InvalidArgument,
        "import targets a sheet or one cell of a sheet",
    )
    .with_format(FileFormat::Ods)
    .with_suggestion("use a path such as /Sheet1 or /Sheet1/B2")
}

#[cfg(test)]
mod tests {
    use super::date_value;

    #[test]
    fn dates_are_odf_lexical_values() {
        assert_eq!(date_value(0, 0.0, false).as_deref(), Some("1970-01-01"));
        assert_eq!(
            date_value(20_375, 0.0, false).as_deref(),
            Some("2025-10-14")
        );
        assert_eq!(
            date_value(20_375, 30_600.5, true).as_deref(),
            Some("2025-10-14T08:30:00.5")
        );
        assert_eq!(date_value(-719_469, 0.0, false), None);
    }
}
