//! SYLK (SYmbolic LinK), the Multiplan text spreadsheet format Excel reads.
//!
//! A SYLK file is a sequence of records, one per line: `ID` identifies the
//! producer, `B` the bounds, `C` a cell (`C;Y<row>;X<column>;K<value>` or
//! `;E<formula>`, whose `K` is then the cached value) and `E` ends the
//! file. The reader keeps cell values and formulas; other records are
//! skipped. Coordinates outside the
//! spreadsheet grid are [`Error::OutOfGrid`](crate::Error::OutOfGrid).
//!
//! ```
//! use litchi_core::sheet::CellValue;
//! use litchi_sheet_text::sylk::{self, Options};
//!
//! let rows = sylk::read(&b"ID;PWXL\nC;Y1;X2;K\"total\"\nC;Y2;X2;K42\nE\n"[..], &Options::default())?;
//! assert_eq!(rows[0], [CellValue::Empty, CellValue::String("total".into())]);
//! assert_eq!(rows[1][1], CellValue::Int(42));
//! # Ok::<(), litchi_sheet_text::Error>(())
//! ```

use std::io::{Read, Write};

use litchi_core::sheet::CellValue;

use crate::common::{Common, common_setters, emit, place};
use crate::error::{Format, Result};
use crate::limits::CellBudget;

/// SYLK options: the source encoding, the byte-order mark to write and the
/// read limits.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Options {
    common: Common,
}

impl Options {
    common_setters!();
}

/// Read the cells of a SYLK source.
///
/// # Errors
///
/// [`Error::Io`](crate::Error::Io), [`Error::Encoding`](crate::Error::Encoding),
/// [`Error::OutOfGrid`](crate::Error::OutOfGrid) and
/// [`Error::Limit`](crate::Error::Limit).
pub fn read(reader: impl Read, options: &Options) -> Result<Vec<Vec<CellValue>>> {
    options.common.read(reader, parse)
}

fn parse(text: &str, budget: &mut CellBudget) -> Result<Vec<Vec<CellValue>>> {
    let mut rows = Vec::new();
    for (index, line) in text.lines().enumerate() {
        let record = line.trim();
        if let Some(fields) = record.strip_prefix("C;") {
            cell(fields, &mut rows, budget, index + 1)?;
        }
    }
    Ok(rows)
}

fn cell(
    fields: &str,
    rows: &mut Vec<Vec<CellValue>>,
    budget: &mut CellBudget,
    line: usize,
) -> Result<()> {
    let mut row: Option<usize> = None;
    let mut column: Option<usize> = None;
    let mut value: Option<String> = None;
    let mut formula: Option<String> = None;
    for field in split_fields(fields) {
        let mut characters = field.chars();
        let Some(prefix) = characters.next() else {
            continue;
        };
        let content = characters.as_str();
        match prefix {
            'Y' => {
                row = content
                    .parse::<usize>()
                    .ok()
                    .map(|row| row.saturating_sub(1))
            },
            'X' => {
                column = content
                    .parse::<usize>()
                    .ok()
                    .map(|column| column.saturating_sub(1))
            },
            'K' => value = Some(constant(content)),
            'E' => formula = Some(content.to_owned()),
            _ => {},
        }
    }
    let (Some(row), Some(column)) = (row, column) else {
        return Ok(());
    };
    place(rows, row, column, budget, Format::Sylk, line)?;
    let value = value.map(|value| CellValue::infer_from_str(&value));
    rows[row][column] = match (formula, value) {
        (Some(formula), cached) => CellValue::Formula {
            formula,
            cached_value: cached.map(Box::new),
            is_array: false,
            array_range: None,
        },
        (None, Some(value)) => value,
        (None, None) => return Ok(()),
    };
    Ok(())
}

/// A `K` constant: a quoted string (doubled quotes inside) or a bare
/// number or boolean.
fn constant(content: &str) -> String {
    match content
        .strip_prefix('"')
        .map(|rest| rest.strip_suffix('"').unwrap_or(rest))
    {
        Some(quoted) => quoted.replace("\"\"", "\""),
        None => content.to_owned(),
    }
}

/// Split a record's fields at single `;`; a doubled `;;` is a literal
/// semicolon inside a field.
fn split_fields(record: &str) -> Vec<std::borrow::Cow<'_, str>> {
    let mut fields = Vec::new();
    let bytes = record.as_bytes();
    let mut start = 0;
    let mut index = 0;
    let mut escaped = false;
    while index < bytes.len() {
        if bytes[index] == b';' {
            if bytes.get(index + 1) == Some(&b';') {
                escaped = true;
                index += 2;
                continue;
            }
            fields.push(unescape(&record[start..index], escaped));
            escaped = false;
            start = index + 1;
        }
        index += 1;
    }
    fields.push(unescape(&record[start..], escaped));
    fields
}

fn unescape(field: &str, escaped: bool) -> std::borrow::Cow<'_, str> {
    if escaped {
        std::borrow::Cow::Owned(field.replace(";;", ";"))
    } else {
        std::borrow::Cow::Borrowed(field)
    }
}

/// Write `rows` as SYLK: an `ID` record, the bounds, one `C` record per
/// non-empty cell and the `E` record.
///
/// # Errors
///
/// [`Error::Io`](crate::Error::Io).
pub fn write(rows: &[Vec<CellValue>], writer: impl Write, options: &Options) -> Result<()> {
    emit(writer, options.common.write_bom, |writer| {
        body(rows, writer)
    })
}

fn body(rows: &[Vec<CellValue>], writer: &mut dyn Write) -> Result<()> {
    writeln!(writer, "ID;PWXL;N;E")?;
    let columns = rows.iter().map(Vec::len).max().unwrap_or(0);
    if !rows.is_empty() && columns > 0 {
        writeln!(writer, "B;Y{};X{}", rows.len(), columns)?;
    }
    for (row_index, cells) in rows.iter().enumerate() {
        let row = row_index + 1;
        for (column_index, cell) in cells.iter().enumerate() {
            let column = column_index + 1;
            match cell {
                CellValue::Empty => {},
                CellValue::Int(value) => writeln!(writer, "C;Y{row};X{column};K{value}")?,
                CellValue::Float(value) => writeln!(writer, "C;Y{row};X{column};K{value}")?,
                CellValue::Bool(value) => writeln!(
                    writer,
                    "C;Y{row};X{column};K\"{}\"",
                    if *value { "TRUE" } else { "FALSE" }
                )?,
                CellValue::String(value) => writeln!(
                    writer,
                    "C;Y{row};X{column};K\"{}\"",
                    value.replace('"', "\"\"").replace(';', ";;")
                )?,
                CellValue::DateTime(value) => writeln!(writer, "C;Y{row};X{column};K{value}")?,
                CellValue::Error(value) => writeln!(writer, "C;Y{row};X{column};K\"{value}\"")?,
                CellValue::Formula { formula, .. } => {
                    writeln!(writer, "C;Y{row};X{column};E{}", formula.replace(';', ";;"))?;
                },
            }
        }
    }
    writeln!(writer, "E")?;
    Ok(())
}
