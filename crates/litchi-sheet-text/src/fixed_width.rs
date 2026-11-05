//! Fixed-width text columns (space-aligned PRN).
//!
//! Widths count characters. The reader either takes caller widths
//! ([`Options::column_widths`]) or detects them: a column boundary falls in
//! the middle of each run of character positions that are whitespace in at
//! least half the non-blank lines. Blank lines are skipped.
//!
//! ```
//! use litchi_core::sheet::CellValue;
//! use litchi_sheet_text::fixed_width::{self, Options};
//!
//! let options = Options::default().column_widths(vec![6, 3]);
//! let rows = fixed_width::read(&b"Ada   36\nGrace 85\n"[..], &options)?;
//! assert_eq!(rows[1], [CellValue::String("Grace".into()), CellValue::Int(85)]);
//! # Ok::<(), litchi_sheet_text::Error>(())
//! ```

use std::io::{Read, Write};

use litchi_core::sheet::CellValue;

use crate::common::{Common, common_setters, emit};
use crate::error::Result;
use crate::limits::CellBudget;

/// Fixed-width options: the column widths (or width detection), field
/// trimming, the source encoding, the byte-order mark to write and the read
/// limits.
///
/// The default detects widths when reading, computes them from the content
/// when writing, and trims fields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Options {
    common: Common,
    widths: Vec<usize>,
    detect: bool,
    trim: bool,
}

impl Default for Options {
    fn default() -> Self {
        Self {
            common: Common::default(),
            widths: Vec::new(),
            detect: true,
            trim: true,
        }
    }
}

impl Options {
    common_setters!();

    /// Use these column widths, in characters, for reading and writing, and
    /// stop detecting widths.
    #[must_use]
    pub fn column_widths(mut self, widths: Vec<usize>) -> Self {
        self.widths = widths;
        self.detect = false;
        self
    }

    /// Whether the reader detects column widths.
    #[must_use]
    pub const fn detect_widths(mut self, detect: bool) -> Self {
        self.detect = detect;
        self
    }

    /// Whether whitespace is trimmed from both ends of each field.
    #[must_use]
    pub const fn trim(mut self, trim: bool) -> Self {
        self.trim = trim;
        self
    }
}

/// Read the rows of a fixed-width source.
///
/// # Errors
///
/// [`Error::Io`](crate::Error::Io), [`Error::Encoding`](crate::Error::Encoding)
/// and [`Error::Limit`](crate::Error::Limit).
pub fn read(reader: impl Read, options: &Options) -> Result<Vec<Vec<CellValue>>> {
    options
        .common
        .read(reader, |text, budget| parse(text, options, budget))
}

fn parse(text: &str, options: &Options, budget: &mut CellBudget) -> Result<Vec<Vec<CellValue>>> {
    let lines: Vec<&str> = text
        .lines()
        .filter(|line| !line.trim().is_empty())
        .collect();
    if lines.is_empty() {
        return Ok(Vec::new());
    }
    let widths = if options.detect {
        detect(&lines)?
    } else {
        options.widths.clone()
    };
    let mut rows = Vec::new();
    for line in lines {
        let mut row = Vec::new();
        let mut rest = line;
        for &width in &widths {
            if rest.is_empty() {
                break;
            }
            let end = rest
                .char_indices()
                .nth(width)
                .map_or(rest.len(), |(offset, _)| offset);
            let (field, tail) = rest.split_at(end);
            rest = tail;
            let field = if options.trim { field.trim() } else { field };
            budget.charge(1)?;
            row.try_reserve(1)?;
            row.push(CellValue::infer_from_str(field));
        }
        if !row.is_empty() {
            rows.try_reserve(1)?;
            rows.push(row);
        }
    }
    Ok(rows)
}

/// Detect column widths from the whitespace positions of `lines`.
fn detect(lines: &[&str]) -> Result<Vec<usize>> {
    let longest = lines
        .iter()
        .map(|line| line.chars().count())
        .max()
        .unwrap_or(0);
    let mut spaces = Vec::new();
    spaces.try_reserve(longest)?;
    spaces.resize(longest, 0usize);
    for line in lines {
        for (position, character) in line.chars().enumerate() {
            if character.is_whitespace() {
                spaces[position] += 1;
            }
        }
    }
    let threshold = lines.len() / 2;
    let mut boundaries = Vec::new();
    let mut run_start = None;
    for (position, &count) in spaces.iter().enumerate() {
        match run_start {
            None if count >= threshold => run_start = Some(position),
            Some(start) if count < threshold => {
                boundaries.push((start + position) / 2);
                run_start = None;
            },
            _ => {},
        }
    }
    if boundaries.is_empty() {
        return Ok(vec![longest]);
    }
    let mut widths = Vec::with_capacity(boundaries.len() + 1);
    let mut previous = 0;
    for boundary in boundaries {
        if boundary > previous {
            widths.push(boundary - previous);
            previous = boundary;
        }
    }
    widths.push(longest.saturating_sub(previous).max(1));
    Ok(widths)
}

/// Write `rows` left-aligned in fixed-width columns, each line's trailing
/// spaces removed. Without caller widths, each column is two characters
/// wider than its longest value; a column beyond the caller widths is ten
/// characters wide. A longer value is cut to its column.
///
/// # Errors
///
/// [`Error::Io`](crate::Error::Io).
pub fn write(rows: &[Vec<CellValue>], writer: impl Write, options: &Options) -> Result<()> {
    emit(writer, options.common.write_bom, |writer| {
        body(rows, writer, options)
    })
}

fn body(rows: &[Vec<CellValue>], writer: &mut dyn Write, options: &Options) -> Result<()> {
    let widths = if options.widths.is_empty() {
        content_widths(rows)
    } else {
        options.widths.clone()
    };
    for cells in rows {
        let mut line = String::new();
        for (column, cell) in cells.iter().enumerate() {
            let width = widths.get(column).copied().unwrap_or(10);
            let text = render(cell);
            let mut written = 0;
            for character in text.chars().take(width) {
                line.push(character);
                written += 1;
            }
            line.extend(std::iter::repeat_n(' ', width - written));
        }
        writeln!(writer, "{}", line.trim_end())?;
    }
    Ok(())
}

fn content_widths(rows: &[Vec<CellValue>]) -> Vec<usize> {
    let columns = rows.iter().map(Vec::len).max().unwrap_or(0);
    let mut widths = vec![0usize; columns];
    for cells in rows {
        for (column, cell) in cells.iter().enumerate() {
            widths[column] = widths[column].max(render(cell).chars().count() + 2);
        }
    }
    widths
}

fn render(cell: &CellValue) -> String {
    match cell {
        CellValue::Empty => String::new(),
        CellValue::Bool(value) => if *value { "TRUE" } else { "FALSE" }.to_owned(),
        CellValue::Int(value) => value.to_string(),
        CellValue::Float(value) | CellValue::DateTime(value) => value.to_string(),
        CellValue::String(value) | CellValue::Error(value) => value.clone(),
        CellValue::Formula {
            cached_value,
            formula,
            ..
        } => match cached_value {
            Some(cached) => render(cached),
            None => format!("={formula}"),
        },
    }
}
