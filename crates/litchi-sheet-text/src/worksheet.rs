//! The worksheet and cell views of a text [`Workbook`](crate::Workbook).

use std::borrow::Cow;

use litchi_core::sheet::{Cell as CellTrait, CellValue, Worksheet as WorksheetTrait};

use crate::error::{Error, Result};

/// One cell: a one-based position and its value.
#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    row: u32,
    column: u32,
    value: CellValue,
}

impl Cell {
    /// A cell at the one-based `row` and `column`.
    #[must_use]
    pub const fn new(row: u32, column: u32, value: CellValue) -> Self {
        Self { row, column, value }
    }
}

impl CellTrait for Cell {
    fn row(&self) -> u32 {
        self.row
    }

    fn column(&self) -> u32 {
        self.column
    }

    fn coordinate(&self) -> String {
        let mut letters = Vec::new();
        let mut column = self.column;
        while column > 0 {
            column -= 1;
            letters.push(b'A' + (column % 26) as u8);
            column /= 26;
        }
        letters.reverse();
        let mut coordinate = String::from_utf8(letters).unwrap_or_default();
        coordinate.push_str(&self.row.to_string());
        coordinate
    }

    fn value(&self) -> &CellValue {
        &self.value
    }
}

/// The one worksheet of a text [`Workbook`](crate::Workbook), borrowing its
/// rows.
#[derive(Debug, Clone, Copy)]
pub struct Worksheet<'a> {
    rows: &'a [Vec<CellValue>],
    name: &'a str,
}

impl<'a> Worksheet<'a> {
    pub(crate) const fn new(rows: &'a [Vec<CellValue>], name: &'a str) -> Self {
        Self { rows, name }
    }

    /// The value at the one-based `row` and `column`; empty outside the
    /// rows.
    fn value(&self, row: u32, column: u32) -> &'a CellValue {
        let (Some(row), Some(column)) = (row.checked_sub(1), column.checked_sub(1)) else {
            return CellValue::EMPTY;
        };
        self.rows
            .get(row as usize)
            .and_then(|cells| cells.get(column as usize))
            .unwrap_or(CellValue::EMPTY)
    }
}

/// Parse an A1 coordinate into one-based `(row, column)`.
fn coordinate(text: &str) -> Option<(u32, u32)> {
    let split = text
        .bytes()
        .position(|byte| byte.is_ascii_digit())
        .unwrap_or(text.len());
    let (letters, digits) = text.split_at(split);
    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    let mut column = 0u32;
    for byte in letters.bytes() {
        if !byte.is_ascii_alphabetic() {
            return None;
        }
        column = column
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))?;
    }
    let row = digits.parse::<u32>().ok()?;
    Some((row, column))
}

impl<'w> WorksheetTrait for Worksheet<'w> {
    type Error = Error;
    type Cell<'a>
        = Cell
    where
        Self: 'a;
    type Cells<'a>
        = Cells<'a>
    where
        Self: 'a;
    type Rows<'a>
        = Rows<'a>
    where
        Self: 'a;

    fn name(&self) -> &str {
        self.name
    }

    fn row_count(&self) -> usize {
        self.rows.len()
    }

    fn column_count(&self) -> usize {
        self.rows.iter().map(Vec::len).max().unwrap_or(0)
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        if self.rows.is_empty() {
            return None;
        }
        let rows = u32::try_from(self.rows.len()).unwrap_or(u32::MAX);
        let columns = u32::try_from(self.column_count()).unwrap_or(u32::MAX);
        Some((1, 1, rows, columns))
    }

    fn cell(&self, row: u32, column: u32) -> Result<Cell> {
        Ok(Cell::new(row, column, self.value(row, column).clone()))
    }

    fn cell_by_coordinate(&self, text: &str) -> Result<Cell> {
        let (row, column) = coordinate(text).ok_or(Error::InvalidCoordinate)?;
        self.cell(row, column)
    }

    fn cells(&self) -> Cells<'_> {
        Cells {
            rows: self.rows,
            row: 0,
            column: 0,
        }
    }

    fn rows(&self) -> Rows<'_> {
        Rows {
            rows: self.rows.iter(),
        }
    }

    fn row(&self, row: usize) -> Result<Cow<'_, [CellValue]>> {
        self.rows
            .get(row)
            .map(|cells| Cow::Borrowed(cells.as_slice()))
            .ok_or(Error::RowOutOfRange { row })
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        Ok(Cow::Borrowed(self.value(row, column)))
    }
}

/// The single-worksheet iterator.
#[derive(Debug)]
pub struct Worksheets<'a> {
    worksheet: Option<Worksheet<'a>>,
}

impl<'a> Worksheets<'a> {
    pub(crate) const fn new(worksheet: Worksheet<'a>) -> Self {
        Self {
            worksheet: Some(worksheet),
        }
    }
}

impl<'a> Iterator for Worksheets<'a> {
    type Item = Result<Worksheet<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        self.worksheet.take().map(Ok)
    }
}

/// Every stored cell, row by row, one-based.
#[derive(Debug)]
pub struct Cells<'a> {
    rows: &'a [Vec<CellValue>],
    row: usize,
    column: usize,
}

impl Iterator for Cells<'_> {
    type Item = Result<Cell>;

    fn next(&mut self) -> Option<Self::Item> {
        while let Some(cells) = self.rows.get(self.row) {
            if let Some(value) = cells.get(self.column) {
                let cell = Cell::new(
                    u32::try_from(self.row + 1).unwrap_or(u32::MAX),
                    u32::try_from(self.column + 1).unwrap_or(u32::MAX),
                    value.clone(),
                );
                self.column += 1;
                return Some(Ok(cell));
            }
            self.row += 1;
            self.column = 0;
        }
        None
    }
}

/// Every row, borrowed.
#[derive(Debug)]
pub struct Rows<'a> {
    rows: std::slice::Iter<'a, Vec<CellValue>>,
}

impl<'a> Iterator for Rows<'a> {
    type Item = Result<Cow<'a, [CellValue]>>;

    fn next(&mut self) -> Option<Self::Item> {
        self.rows
            .next()
            .map(|cells| Ok(Cow::Borrowed(cells.as_slice())))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn coordinates_follow_a1_letters() {
        assert_eq!(Cell::new(1, 1, CellValue::Empty).coordinate(), "A1");
        assert_eq!(Cell::new(10, 5, CellValue::Empty).coordinate(), "E10");
        assert_eq!(Cell::new(100, 27, CellValue::Empty).coordinate(), "AA100");
        assert_eq!(coordinate("aa100"), Some((100, 27)));
        assert_eq!(coordinate("A"), None);
        assert_eq!(coordinate("1A"), None);
        assert_eq!(coordinate("ZZZZZZZZZZ1"), None);
    }
}
