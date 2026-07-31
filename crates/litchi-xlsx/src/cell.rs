//! Exact cell states and sparse borrowed traversal.

use std::fmt;
use std::ops::Deref;
use std::sync::Arc;

use litchi_sheet::{Cell as Address, Column, Rect};

use crate::error::{Result, invalid};
use crate::formula::Formula;

/// The stored state of one cell record.
///
/// Absence is represented by `Option<Cell>` at lookup sites, so it cannot be
/// confused with an explicitly stored empty cell.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Cell {
    /// A `<c>` record exists but has no primary payload.
    Empty,
    /// An exact non-formula value.
    Value(Value),
    /// A formula and its optional separately qualified cache.
    Formula(Formula),
    /// A cell representation not yet understood by this facade. Core inert
    /// fields are available on [`Unknown`]; the snapshot retains original part
    /// bytes for lossless future saves.
    Unknown(Unknown),
}

/// Exact value stored by SpreadsheetML.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Value {
    Bool(bool),
    /// Numeric lexical form, retained without format-based coercion.
    Number(Number),
    Text(Text),
    /// ISO 8601 lexical form from a `t="d"` cell.
    Date(Text),
    Error(ErrorValue),
}

/// An exact SpreadsheetML number.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct Number(Box<str>);

impl Number {
    /// Validate and retain a numeric lexical form without normalizing it.
    pub fn new(value: impl Into<Box<str>>) -> Result<Self> {
        let value = value.into();
        let parsed = value
            .trim()
            .parse::<f64>()
            .map_err(|_| invalid(format!("invalid worksheet number '{value}'")))?;
        if !parsed.is_finite() {
            return Err(invalid(format!("non-finite worksheet number '{value}'")));
        }
        Ok(Self(value))
    }

    /// Exact stored lexical form.
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Convert the value to IEEE-754 binary64.
    ///
    /// `None` keeps this accessor safe if a future lossless reader accepts a
    /// numeric lexical form outside Rust's binary64 parser.
    pub fn as_f64(&self) -> Option<f64> {
        self.0
            .trim()
            .parse()
            .ok()
            .filter(|value: &f64| value.is_finite())
    }
}

impl fmt::Debug for Number {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_tuple("Number").field(&self.0).finish()
    }
}

impl fmt::Display for Number {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

/// Cheaply cloned immutable text, including resolved shared strings.
#[derive(Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Text(Arc<str>);

impl Text {
    /// Move or borrow text into an immutable value.
    pub fn new(value: impl Into<Arc<str>>) -> Self {
        Self(value.into())
    }

    /// Borrow the text.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl Deref for Text {
    type Target = str;

    fn deref(&self) -> &Self::Target {
        self.as_str()
    }
}

impl AsRef<str> for Text {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl From<String> for Text {
    fn from(value: String) -> Self {
        Self(value.into())
    }
}

impl From<&str> for Text {
    fn from(value: &str) -> Self {
        Self(value.into())
    }
}

impl fmt::Debug for Text {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl fmt::Display for Text {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

/// Spreadsheet error value.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ErrorValue {
    Null,
    DivZero,
    Value,
    Ref,
    Name,
    Num,
    NotAvailable,
    GettingData,
    Spill,
    Calc,
    Field,
    Blocked,
    Connect,
    Unknown(Text),
}

impl ErrorValue {
    pub(crate) fn parse(value: &str) -> Self {
        match value.trim() {
            "#NULL!" => Self::Null,
            "#DIV/0!" => Self::DivZero,
            "#VALUE!" => Self::Value,
            "#REF!" => Self::Ref,
            "#NAME?" => Self::Name,
            "#NUM!" => Self::Num,
            "#N/A" => Self::NotAvailable,
            "#GETTING_DATA" => Self::GettingData,
            "#SPILL!" => Self::Spill,
            "#CALC!" => Self::Calc,
            "#FIELD!" => Self::Field,
            "#BLOCKED!" => Self::Blocked,
            "#CONNECT!" => Self::Connect,
            other => Self::Unknown(other.into()),
        }
    }
}

/// Bounded diagnostic for a cell encoding not yet modeled semantically.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Unknown {
    kind: Text,
    value: Option<Text>,
    formula: Option<Text>,
}

impl Unknown {
    pub(crate) fn new(
        kind: impl Into<Text>,
        value: Option<String>,
        formula: Option<String>,
    ) -> Self {
        Self {
            kind: kind.into(),
            value: value.map(Text::from),
            formula: formula.map(Text::from),
        }
    }

    /// Producer cell type or formula form that was not recognized.
    pub fn kind(&self) -> &str {
        &self.kind
    }

    /// Uninterpreted value text, when present.
    pub fn value(&self) -> Option<&str> {
        self.value.as_deref()
    }

    /// Uninterpreted formula text, when present.
    pub fn formula(&self) -> Option<&str> {
        self.formula.as_deref()
    }
}

#[derive(Debug)]
pub(crate) struct Stored {
    pub(crate) address: Address,
    pub(crate) cell: Cell,
    // Retained for the shared-style facade. Native indexes never escape this
    // migration boundary.
    #[allow(dead_code)]
    pub(crate) style: Option<u32>,
    #[allow(dead_code)]
    pub(crate) cell_metadata: Option<u32>,
    #[allow(dead_code)]
    pub(crate) value_metadata: Option<u32>,
}

#[derive(Debug, Default)]
pub(crate) struct Store {
    cells: Box<[Stored]>,
    extent: Option<Rect>,
}

impl Store {
    pub(crate) fn from_unsorted(mut cells: Vec<Stored>) -> Result<Self> {
        cells.sort_unstable_by_key(|entry| entry.address);
        if let Some(pair) = cells
            .windows(2)
            .find(|pair| pair[0].address == pair[1].address)
        {
            return Err(invalid(format!(
                "duplicate worksheet cell at {:?}",
                pair[0].address
            )));
        }

        let extent = extent(&cells)?;
        Ok(Self {
            cells: cells.into_boxed_slice(),
            extent,
        })
    }

    pub(crate) fn get(&self, address: Address) -> Option<&Cell> {
        self.cells
            .binary_search_by_key(&address, |entry| entry.address)
            .ok()
            .and_then(|index| self.cells.get(index))
            .map(|entry| &entry.cell)
    }

    pub(crate) fn cells(&self, range: Rect) -> Cells<'_> {
        let first = Address::new(range.start().row(), Column::FIRST);
        let start = self.cells.partition_point(|entry| entry.address < first);
        Cells {
            remaining: &self.cells[start..],
            range,
        }
    }

    pub(crate) fn extent(&self) -> Option<Rect> {
        self.extent
    }
}

fn extent(cells: &[Stored]) -> Result<Option<Rect>> {
    let Some(first) = cells.first() else {
        return Ok(None);
    };
    let mut min_row = first.address.row().get();
    let mut min_column = first.address.column().get();
    let mut max_row = min_row;
    let mut max_column = min_column;
    for entry in &cells[1..] {
        let row = entry.address.row().get();
        let column = entry.address.column().get();
        min_row = min_row.min(row);
        min_column = min_column.min(column);
        max_row = max_row.max(row);
        max_column = max_column.max(column);
    }
    let start = Address::at(min_row, min_column)?;
    Rect::new(start, max_row + 1, max_column + 1)
        .map(Some)
        .map_err(|error| invalid(error.to_string()))
}

/// Borrowed sparse cells inside a half-open range.
#[derive(Debug)]
pub struct Cells<'a> {
    remaining: &'a [Stored],
    range: Rect,
}

impl<'a> Iterator for Cells<'a> {
    type Item = (Address, &'a Cell);

    fn next(&mut self) -> Option<Self::Item> {
        while let Some((entry, remaining)) = self.remaining.split_first() {
            self.remaining = remaining;
            if entry.address.row().get() >= self.range.end().0 {
                self.remaining = &[];
                return None;
            }
            if self.range.contains(entry.address) {
                return Some((entry.address, &entry.cell));
            }
        }
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn numbers_preserve_lexemes_and_convert_explicitly() {
        let number = Number::new("  -0.000  ").expect("valid number");
        assert_eq!(number.as_str(), "  -0.000  ");
        assert_eq!(number.as_f64(), Some(-0.0));
        assert!(Number::new("NaN").is_err());
        assert!(Number::new("not a number").is_err());
    }
}
