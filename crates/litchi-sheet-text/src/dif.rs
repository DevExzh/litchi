//! DIF (Data Interchange Format), the VisiCalc interchange format.
//!
//! A DIF file has a header of `TABLE`, `VECTORS` (columns), `TUPLES`
//! (rows) and other sections, then `DATA`. Each data value is a
//! `<type>,<number>` line followed by a string line: `-1,0` with `BOT`
//! starts a row, `0,<number>` is a number, and `1,0` is a string (or the
//! booleans `TRUE`/`FALSE`, and `NA` or `""` for an empty cell). A file
//! without a `DATA` section is [`Error::Syntax`].
//!
//! ```
//! use litchi_core::sheet::CellValue;
//! use litchi_sheet_text::dif::{self, Options};
//!
//! let mut out = Vec::new();
//! dif::write(&[vec![CellValue::String("a".into()), CellValue::Int(2)]], &mut out, &Options::default())?;
//! let rows = dif::read(&out[..], &Options::default())?;
//! assert_eq!(rows, [vec![CellValue::String("a".into()), CellValue::Float(2.0)]]);
//! # Ok::<(), litchi_sheet_text::Error>(())
//! ```

use std::io::{Read, Write};

use litchi_core::sheet::CellValue;

use crate::common::{Common, common_setters, emit};
use crate::error::{Error, Format, Result};
use crate::limits::CellBudget;

/// DIF options: the source encoding, the byte-order mark to write and the
/// read limits.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Options {
    common: Common,
}

impl Options {
    common_setters!();
}

/// Read the rows of a DIF source.
///
/// # Errors
///
/// [`Error::Io`], [`Error::Encoding`], [`Error::Syntax`] and
/// [`Error::Limit`].
pub fn read(reader: impl Read, options: &Options) -> Result<Vec<Vec<CellValue>>> {
    options.common.read(reader, parse)
}

/// A `<type>,<number>` line.
fn value_line(line: &str) -> Option<(i32, f64)> {
    let mut parts = line.split(',');
    let kind = parts.next()?.trim().parse::<i32>().ok()?;
    let number = parts.next()?.trim().parse::<f64>().unwrap_or(0.0);
    Some((kind, number))
}

/// Lines with a running count.
struct Counted<'a> {
    lines: std::str::Lines<'a>,
    count: usize,
}

impl<'a> Counted<'a> {
    fn next(&mut self) -> Option<&'a str> {
        let line = self.lines.next();
        self.count += usize::from(line.is_some());
        line
    }
}

fn parse(text: &str, budget: &mut CellBudget) -> Result<Vec<Vec<CellValue>>> {
    let mut lines = Counted {
        lines: text.lines(),
        count: 0,
    };
    // The header: skip every section up to `DATA` and its two lines.
    loop {
        let Some(line) = lines.next() else {
            return Err(Error::Syntax {
                format: Format::Dif,
                line: lines.count,
            });
        };
        if line.trim() == "DATA" {
            lines.next();
            lines.next();
            break;
        }
    }
    let mut rows = Vec::new();
    let mut row: Vec<CellValue> = Vec::new();
    while let Some(line) = lines.next() {
        let Some((kind, number)) = value_line(line.trim()) else {
            continue;
        };
        let value = match kind {
            -1 if number == 0.0 => {
                if !row.is_empty() {
                    rows.try_reserve(1)?;
                    rows.push(std::mem::take(&mut row));
                }
                continue;
            },
            // The indicator line (`V`, `NA`, ...) after the number has no
            // comma and is skipped as it comes.
            0 => CellValue::Float(number),
            1 => {
                let text = lines.next().unwrap_or_default().trim().trim_matches('"');
                if number == 0.0 {
                    match text.to_ascii_uppercase().as_str() {
                        "TRUE" | "V" => CellValue::Bool(true),
                        "FALSE" => CellValue::Bool(false),
                        "NA" | "" => CellValue::Empty,
                        _ => CellValue::String(text.to_owned()),
                    }
                } else {
                    CellValue::String(text.to_owned())
                }
            },
            _ => continue,
        };
        budget.charge(1)?;
        row.try_reserve(1)?;
        row.push(value);
    }
    if !row.is_empty() {
        rows.try_reserve(1)?;
        rows.push(row);
    }
    Ok(rows)
}

/// Write `rows` as DIF. Empty cells are `NA`; a formula writes its cached
/// number, or the string `0` without one.
///
/// # Errors
///
/// [`Error::Io`].
pub fn write(rows: &[Vec<CellValue>], writer: impl Write, options: &Options) -> Result<()> {
    emit(writer, options.common.write_bom, |writer| {
        body(rows, writer)
    })
}

fn body(rows: &[Vec<CellValue>], writer: &mut dyn Write) -> Result<()> {
    let tuples = rows.len();
    let vectors = rows.iter().map(Vec::len).max().unwrap_or(0);
    write!(
        writer,
        "TABLE\n0,1\n\"LITCHI\"\nVECTORS\n0,{vectors}\n\"\"\nTUPLES\n0,{tuples}\n\"\"\nDATA\n0,0\n\"\"\n"
    )?;
    for cells in rows {
        writeln!(writer, "-1,0\nBOT")?;
        for cell in cells {
            write_cell(writer, cell)?;
        }
    }
    writeln!(writer, "-1,0\nEOD")?;
    Ok(())
}

fn write_cell(writer: &mut dyn Write, cell: &CellValue) -> Result<()> {
    match cell {
        CellValue::Empty => writeln!(writer, "1,0\n\"NA\"")?,
        CellValue::Int(value) => writeln!(writer, "0,{value}\nV")?,
        CellValue::Float(value) | CellValue::DateTime(value) => writeln!(writer, "0,{value}\nV")?,
        CellValue::Bool(value) => {
            writeln!(writer, "1,0\n\"{}\"", if *value { "TRUE" } else { "FALSE" })?;
        },
        CellValue::String(value) => writeln!(writer, "1,0\n\"{}\"", value.replace('"', "\"\""))?,
        CellValue::Error(value) => writeln!(writer, "1,0\n\"{value}\"")?,
        CellValue::Formula { cached_value, .. } => match cached_value.as_deref() {
            Some(CellValue::Int(value)) => writeln!(writer, "0,{value}\nV")?,
            Some(CellValue::Float(value)) => writeln!(writer, "0,{value}\nV")?,
            _ => writeln!(writer, "1,0\n\"0\"")?,
        },
    }
    Ok(())
}
