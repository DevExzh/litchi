//! Exact cell states and sparse borrowed traversal.

use std::fmt;
use std::hash::{Hash, Hasher};
use std::ops::Deref;
use std::sync::Arc;

use chrono::{DateTime, NaiveDate, NaiveDateTime};
use litchi_sheet::{Cell as Address, Column as ColumnIndex, Rect, Row as RowIndex};

use crate::column;
use crate::error::{Result, allocation, invalid};
use crate::formula::{Formula, Kind};
use crate::layout::Defaults;
use crate::merge;
use crate::row;

const MAX_CELL_CHARACTERS: usize = 32_767;

/// Workbook-local lineage for opaque shared-string identities.
#[derive(Debug)]
pub(crate) struct SharedStringLineage;

/// Opaque shared-string identity retained by semantic patch states.
///
/// The physical `SpreadsheetML` table index remains private. Keys are only
/// comparable inside the workbook lineage that produced them.
#[derive(Clone)]
pub struct SharedStringKey {
    raw: usize,
    lineage: Arc<SharedStringLineage>,
}

impl SharedStringKey {
    pub(crate) fn new(raw: usize, lineage: Arc<SharedStringLineage>) -> Self {
        Self { raw, lineage }
    }

    pub(crate) const fn raw(&self) -> usize {
        self.raw
    }

    pub(crate) fn rebind(&mut self, lineage: &Arc<SharedStringLineage>) {
        self.lineage = Arc::clone(lineage);
    }
}

impl PartialEq for SharedStringKey {
    fn eq(&self, other: &Self) -> bool {
        self.raw == other.raw && Arc::ptr_eq(&self.lineage, &other.lineage)
    }
}

impl Eq for SharedStringKey {}

impl Hash for SharedStringKey {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.raw.hash(state);
        Arc::as_ptr(&self.lineage).hash(state);
    }
}

impl fmt::Debug for SharedStringKey {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("SharedStringKey(..)")
    }
}

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

impl Cell {
    /// Parse a checked A1 reference into one-based `(column, row)` numbers.
    ///
    /// This small lexical helper is shared by worksheet package codecs whose
    /// references are not yet materialized as a [`litchi_sheet::Cell`].
    pub fn reference_to_coords(reference: &str) -> Result<(u32, u32)> {
        const MAX_COLUMN: u32 = 16_384;
        const MAX_ROW: u32 = 1_048_576;

        let bytes = reference.as_bytes();
        let column_end = bytes
            .iter()
            .position(u8::is_ascii_digit)
            .ok_or_else(|| invalid(format!("invalid cell reference '{reference}'")))?;
        if column_end == 0 || column_end == bytes.len() {
            return Err(invalid(format!("invalid cell reference '{reference}'")));
        }

        let mut column = 0u32;
        for byte in &bytes[..column_end] {
            if !byte.is_ascii_alphabetic() {
                return Err(invalid(format!(
                    "invalid column in cell reference '{reference}'"
                )));
            }
            let digit = u32::from(byte.to_ascii_uppercase() - b'A' + 1);
            column = column
                .checked_mul(26)
                .and_then(|value| value.checked_add(digit))
                .ok_or_else(|| {
                    invalid(format!("column overflows in cell reference '{reference}'"))
                })?;
        }
        if column > MAX_COLUMN {
            return Err(invalid(format!(
                "column exceeds Excel limits in cell reference '{reference}'"
            )));
        }

        let row = std::str::from_utf8(&bytes[column_end..])
            .ok()
            .and_then(|value| value.parse::<u32>().ok())
            .filter(|value| *value != 0 && *value <= MAX_ROW)
            .ok_or_else(|| {
                invalid(format!(
                    "row exceeds Excel limits in cell reference '{reference}'"
                ))
            })?;
        Ok((column, row))
    }
}

/// Exact semantic state at one logical worksheet coordinate.
///
/// A covered coordinate is reported before any producer-stored follower cell,
/// because the merge anchor owns its visible content. Sparse stored traversal
/// remains available through [`crate::Worksheet::cells`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum View<'a> {
    /// No cell record or covering merge exists at this coordinate.
    Missing,
    /// The coordinate is covered by a merge whose anchor is `range.start()`.
    Covered(Rect),
    /// One physical cell record is stored at this coordinate.
    Stored(&'a Cell),
}

impl<'a> View<'a> {
    /// Borrow the physical cell state when this coordinate owns one.
    #[must_use]
    pub const fn stored(self) -> Option<&'a Cell> {
        match self {
            Self::Stored(cell) => Some(cell),
            Self::Missing | Self::Covered(_) => None,
        }
    }

    /// Covering merged range, if this is a non-anchor coordinate.
    #[must_use]
    pub const fn merge(self) -> Option<Rect> {
        match self {
            Self::Covered(range) => Some(range),
            Self::Missing | Self::Stored(_) => None,
        }
    }

    /// Whether this coordinate has neither a record nor merge coverage.
    #[must_use]
    pub const fn is_missing(self) -> bool {
        matches!(self, Self::Missing)
    }
}

/// Exact value stored by `SpreadsheetML`.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Value {
    Bool(bool),
    /// Numeric lexical form, retained without format-based coercion.
    Number(Number),
    Text(Text),
    /// Checked ISO 8601 lexical form from a `t="d"` cell.
    Date(Date),
    Error(ErrorValue),
}

impl Value {
    /// Construct inert plain text.
    pub fn text(value: impl Into<Text>) -> Self {
        Self::Text(value.into())
    }

    /// Construct an explicitly typed, checked ISO 8601 date lexical value.
    pub fn date(value: impl Into<Text>) -> Result<Self> {
        Date::new(value).map(Self::Date)
    }

    pub(crate) fn validate_for_write(&self) -> Result<()> {
        let text = match self {
            Self::Error(ErrorValue::Unknown(_)) => {
                return Err(invalid(
                    "unrecognized worksheet error values cannot be authored",
                ));
            },
            Self::Text(text) => Some(text),
            Self::Date(date) => Some(&date.0),
            Self::Bool(_) | Self::Number(_) | Self::Error(_) => None,
        };
        if text.is_some_and(|text| text.chars().count() > MAX_CELL_CHARACTERS) {
            return Err(invalid(format!(
                "cell text exceeds {MAX_CELL_CHARACTERS} characters"
            )));
        }
        Ok(())
    }
}

/// A checked ISO 8601 lexical value for a `SpreadsheetML` date cell.
///
/// The original lexical form is retained so a read/write cycle does not
/// silently normalize producer data.
#[derive(Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Date(Text);

impl Date {
    /// Validate and retain an ISO 8601 date or date-time lexical form.
    pub fn new(value: impl Into<Text>) -> Result<Self> {
        let value = value.into();
        let lexical = value.as_str();
        let valid = NaiveDate::parse_from_str(lexical, "%Y-%m-%d").is_ok()
            || NaiveDateTime::parse_from_str(lexical, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
            || DateTime::parse_from_rfc3339(lexical).is_ok();
        if !valid {
            return Err(invalid(format!(
                "invalid ISO 8601 worksheet date '{lexical}'"
            )));
        }
        Ok(Self(value))
    }

    /// Exact stored lexical form.
    #[must_use]
    pub fn as_str(&self) -> &str {
        self.0.as_str()
    }
}

impl Deref for Date {
    type Target = str;

    fn deref(&self) -> &Self::Target {
        self.as_str()
    }
}

impl AsRef<str> for Date {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Debug for Date {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_tuple("Date").field(&self.0).finish()
    }
}

impl fmt::Display for Date {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl From<bool> for Value {
    fn from(value: bool) -> Self {
        Self::Bool(value)
    }
}

impl From<Number> for Value {
    fn from(value: Number) -> Self {
        Self::Number(value)
    }
}

impl From<Text> for Value {
    fn from(value: Text) -> Self {
        Self::Text(value)
    }
}

impl From<String> for Value {
    fn from(value: String) -> Self {
        Self::Text(value.into())
    }
}

impl From<&str> for Value {
    fn from(value: &str) -> Self {
        Self::Text(value.into())
    }
}

macro_rules! exact_integer_value {
    ($($integer:ty),+ $(,)?) => {
        $(
            impl From<$integer> for Value {
                fn from(value: $integer) -> Self {
                    Self::Number(Number(value.to_string().into_boxed_str()))
                }
            }
        )+
    };
}

exact_integer_value!(i8, i16, i32, u8, u16, u32);

/// Primary payload accepted by [`crate::WorksheetEdit::set`].
///
/// Plain strings are always inert text. Formula interpretation requires an
/// explicit checked [`Formula`].
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Content {
    Value(Value),
    Formula(Formula),
}

impl Content {
    pub(crate) fn validate_for_write(&self) -> Result<()> {
        match self {
            Self::Value(value) => value.validate_for_write(),
            Self::Formula(formula)
                if matches!(formula.kind(), Kind::Scalar) && formula.cached().is_none() =>
            {
                Ok(())
            },
            Self::Formula(formula) if formula.cached().is_some() => Err(invalid(
                "writing a stored formula cache requires an explicit cache policy",
            )),
            Self::Formula(_) => Err(invalid(
                "array and data-table formulas require a range-scoped editor",
            )),
        }
    }

    pub(crate) fn as_cell(&self) -> Cell {
        match self {
            Self::Value(value) => Cell::Value(value.clone()),
            Self::Formula(formula) => Cell::Formula(formula.clone()),
        }
    }
}

impl From<Formula> for Content {
    fn from(value: Formula) -> Self {
        Self::Formula(value)
    }
}

impl<T> From<T> for Content
where
    Value: From<T>,
{
    fn from(value: T) -> Self {
        Self::Value(Value::from(value))
    }
}

/// An exact `SpreadsheetML` number.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct Number(Box<str>);

impl Number {
    /// Validate and retain a numeric lexical form without normalizing it.
    pub fn new(value: impl Into<Box<str>>) -> Result<Self> {
        let value = value.into();
        let parsed = value
            .trim()
            .parse::<f64>()
            .map_err(|_source| invalid(format!("invalid worksheet number '{value}'")))?;
        if !parsed.is_finite() {
            return Err(invalid(format!("non-finite worksheet number '{value}'")));
        }
        Ok(Self(value))
    }

    /// Exact stored lexical form.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Convert the value to IEEE-754 binary64.
    ///
    /// `None` keeps this accessor safe if a future lossless reader accepts a
    /// numeric lexical form outside Rust's binary64 parser.
    #[must_use]
    pub fn as_f64(&self) -> Option<f64> {
        self.0
            .trim()
            .parse()
            .ok()
            .filter(|value: &f64| value.is_finite())
    }
}

impl TryFrom<f64> for Number {
    type Error = crate::Error;

    fn try_from(value: f64) -> Result<Self> {
        if !value.is_finite() {
            return Err(invalid("spreadsheet numbers must be finite"));
        }
        Self::new(value.to_string())
    }
}

impl TryFrom<f32> for Number {
    type Error = crate::Error;

    fn try_from(value: f32) -> Result<Self> {
        Self::try_from(f64::from(value))
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
    #[must_use]
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

    /// Spreadsheet lexical form.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Null => "#NULL!",
            Self::DivZero => "#DIV/0!",
            Self::Value => "#VALUE!",
            Self::Ref => "#REF!",
            Self::Name => "#NAME?",
            Self::Num => "#NUM!",
            Self::NotAvailable => "#N/A",
            Self::GettingData => "#GETTING_DATA",
            Self::Spill => "#SPILL!",
            Self::Calc => "#CALC!",
            Self::Field => "#FIELD!",
            Self::Blocked => "#BLOCKED!",
            Self::Connect => "#CONNECT!",
            Self::Unknown(value) => value,
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
    #[must_use]
    pub fn kind(&self) -> &str {
        &self.kind
    }

    /// Uninterpreted value text, when present.
    #[must_use]
    pub fn value(&self) -> Option<&str> {
        self.value.as_deref()
    }

    /// Uninterpreted formula text, when present.
    #[must_use]
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
    pub(crate) style: Option<u32>,
    // Retained so same-workbook transfers can preserve the exact shared-string
    // item, including rich-text runs, without exposing its physical index.
    pub(crate) shared_string: Option<usize>,
    #[allow(
        dead_code,
        reason = "the cached column supports internal sparse-cell indexing"
    )]
    pub(crate) cell_metadata: Option<u32>,
    #[allow(
        dead_code,
        reason = "the cached row supports internal sparse-cell indexing"
    )]
    pub(crate) value_metadata: Option<u32>,
}

#[derive(Debug, Default)]
pub(crate) struct Store {
    cells: Box<[Stored]>,
    cell_rows: Box<[CellRowStart]>,
    rows: Box<[row::Stored]>,
    columns: Box<[column::Stored]>,
    defaults: Option<Defaults>,
    merges: merge::Index,
    extents: Extents,
}

#[derive(Debug)]
struct CellRowStart {
    row: RowIndex,
    start: usize,
}

/// Distinct worksheet cell-bound summaries.
///
/// Except for the producer-declared hint, these ranges describe stored cell
/// records only. Row/column defaults, drawings, merges, and other sheet objects
/// are intentionally not folded into the semantic cell extents.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Extents {
    declared: Option<Rect>,
    stored: Option<Rect>,
    content: Option<Rect>,
    styled: Option<Rect>,
}

impl Extents {
    /// Producer-declared worksheet `dimension`, when present.
    #[must_use]
    pub const fn declared(&self) -> Option<Rect> {
        self.declared
    }

    /// Bounds of every explicit cell record, including empty metadata cells.
    #[must_use]
    pub const fn stored(&self) -> Option<Rect> {
        self.stored
    }

    /// Bounds of cells with a value, formula, or unknown primary payload.
    #[must_use]
    pub const fn content(&self) -> Option<Rect> {
        self.content
    }

    /// Bounds of cells with an explicit local shared-style reference.
    #[must_use]
    pub const fn styled(&self) -> Option<Rect> {
        self.styled
    }

    /// Bounds of cells with content or direct local formatting.
    ///
    /// This does not include formatting inherited from row/column defaults.
    #[must_use]
    pub const fn used(&self) -> Option<Rect> {
        match (self.content, self.styled) {
            (Some(content), Some(styled)) => Some(content.union(styled)),
            (Some(content), None) => Some(content),
            (None, Some(styled)) => Some(styled),
            (None, None) => None,
        }
    }
}

impl Store {
    pub(crate) fn from_unsorted(
        mut cells: Vec<Stored>,
        mut rows: Vec<row::Stored>,
        columns: Box<[column::Stored]>,
        defaults: Option<Defaults>,
        merges: Vec<Rect>,
        declared: Option<Rect>,
    ) -> Result<Self> {
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

        rows.sort_unstable_by_key(|entry| entry.index);
        if let Some(pair) = rows.windows(2).find(|pair| pair[0].index == pair[1].index) {
            return Err(invalid(format!(
                "duplicate worksheet row {}",
                pair[0].index.get() + 1
            )));
        }

        let mut stored = Bounds::default();
        let mut content = Bounds::default();
        let mut styled = Bounds::default();
        let mut cell_rows = Vec::new();
        for (index, entry) in cells.iter().enumerate() {
            if index == 0 || cells[index - 1].address.row() != entry.address.row() {
                cell_rows
                    .try_reserve(1)
                    .map_err(|source| allocation("worksheet cell-row index", source))?;
                cell_rows.push(CellRowStart {
                    row: entry.address.row(),
                    start: index,
                });
            }
            stored.push(entry.address);
            if !matches!(entry.cell, Cell::Empty) {
                content.push(entry.address);
            }
            if entry.style.is_some() {
                styled.push(entry.address);
            }
        }
        let merges = merge::Index::new(merges)?;
        Ok(Self {
            cells: cells.into_boxed_slice(),
            cell_rows: cell_rows.into_boxed_slice(),
            rows: rows.into_boxed_slice(),
            columns,
            defaults,
            merges,
            extents: Extents {
                declared,
                stored: stored.finish()?,
                content: content.finish()?,
                styled: styled.finish()?,
            },
        })
    }

    pub(crate) fn view(&self, address: Address) -> View<'_> {
        if let Some(range) = self.merges.containing(address)
            && range.start() != address
        {
            return View::Covered(range);
        }
        self.entry(address)
            .map_or(View::Missing, |entry| View::Stored(&entry.cell))
    }

    #[cfg(test)]
    pub(crate) fn get(&self, address: Address) -> Option<&Cell> {
        self.entry(address).map(|entry| &entry.cell)
    }

    pub(crate) fn entry(&self, address: Address) -> Option<&Stored> {
        self.cells
            .binary_search_by_key(&address, |entry| entry.address)
            .ok()
            .and_then(|index| self.cells.get(index))
    }

    pub(crate) fn entries(&self) -> &[Stored] {
        &self.cells
    }

    pub(crate) fn row(&self, index: RowIndex) -> row::Row<'_> {
        row::Row::new(index, self.row_entry(index))
    }

    pub(crate) fn row_entry(&self, index: RowIndex) -> Option<&row::Stored> {
        self.rows
            .binary_search_by_key(&index, |entry| entry.index)
            .ok()
            .and_then(|position| self.rows.get(position))
    }

    pub(crate) fn row_entries(&self) -> &[row::Stored] {
        &self.rows
    }

    pub(crate) fn rows(&self) -> row::Rows<'_> {
        row::Rows::new(&self.rows)
    }

    pub(crate) fn column(&self, index: ColumnIndex) -> column::Column<'_> {
        column::Column::new(index, self.column_entry(index))
    }

    pub(crate) fn column_entry(&self, index: ColumnIndex) -> Option<&column::Stored> {
        column::entry(&self.columns, index)
    }

    pub(crate) fn column_entries(&self) -> &[column::Stored] {
        &self.columns
    }

    pub(crate) fn columns(&self) -> column::Columns<'_> {
        column::Columns::new(&self.columns)
    }

    pub(crate) const fn defaults(&self) -> Option<&Defaults> {
        self.defaults.as_ref()
    }

    pub(crate) fn merges(&self) -> merge::Merges<'_> {
        self.merges.iter()
    }

    pub(crate) fn merge_ranges(&self) -> &[Rect] {
        self.merges.as_slice()
    }

    pub(crate) fn cells(&self, range: Rect) -> Cells<'_> {
        let start = self
            .cell_rows
            .partition_point(|entry| entry.row < range.start().row());
        Cells {
            cells: &self.cells,
            rows: &self.cell_rows[start..],
            current: &[],
            range,
        }
    }

    pub(crate) const fn extents(&self) -> &Extents {
        &self.extents
    }
}

#[derive(Debug, Default)]
struct Bounds {
    value: Option<(u32, u32, u32, u32)>,
}

impl Bounds {
    fn push(&mut self, address: Address) {
        let row = address.row().get();
        let column = address.column().get();
        self.value = Some(self.value.map_or(
            (row, column, row, column),
            |(min_row, min_column, max_row, max_column)| {
                (
                    min_row.min(row),
                    min_column.min(column),
                    max_row.max(row),
                    max_column.max(column),
                )
            },
        ));
    }

    fn finish(self) -> Result<Option<Rect>> {
        let Some((min_row, min_column, max_row, max_column)) = self.value else {
            return Ok(None);
        };
        let start = Address::at(min_row, min_column)?;
        Rect::new(start, max_row + 1, max_column + 1)
            .map(Some)
            .map_err(|error| invalid(error.to_string()))
    }
}

/// Borrowed sparse cells inside a half-open range.
#[derive(Debug)]
pub struct Cells<'a> {
    cells: &'a [Stored],
    rows: &'a [CellRowStart],
    current: &'a [Stored],
    range: Rect,
}

impl<'a> Iterator for Cells<'a> {
    type Item = (Address, &'a Cell);

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            if let Some((entry, remaining)) = self.current.split_first() {
                self.current = remaining;
                return Some((entry.address, &entry.cell));
            }

            let (row, remaining_rows) = self.rows.split_first()?;
            self.rows = remaining_rows;
            if row.row.get() >= self.range.end().0 {
                self.rows = &[];
                return None;
            }

            let start_column = self.range.start().column().get();
            let end = remaining_rows
                .first()
                .map_or(self.cells.len(), |next| next.start);
            let row_cells = &self.cells[row.start..end];
            let start = if start_column == ColumnIndex::FIRST.get() {
                0
            } else {
                row_cells.partition_point(|entry| entry.address.column().get() < start_column)
            };
            let selected = &row_cells[start..];
            let selected_len = if self.range.end().1 == ColumnIndex::LAST.get() + 1 {
                selected.len()
            } else if self.range.end().1 == start_column + 1 {
                usize::from(
                    selected
                        .first()
                        .is_some_and(|entry| entry.address.column().get() == start_column),
                )
            } else {
                selected.partition_point(|entry| entry.address.column().get() < self.range.end().1)
            };
            self.current = &selected[..selected_len];
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn empty_store(addresses: &[(u32, u32)]) -> Store {
        let cells = addresses
            .iter()
            .map(|&(row, column)| Stored {
                address: Address::at(row, column).expect("bounded test address"),
                cell: Cell::Empty,
                style: None,
                shared_string: None,
                cell_metadata: None,
                value_metadata: None,
            })
            .collect();
        Store::from_unsorted(cells, Vec::new(), Box::new([]), None, Vec::new(), None)
            .expect("valid cell store")
    }

    #[test]
    fn range_cells_skip_columns_outside_each_selected_row() {
        let store = empty_store(&[
            (0, 0),
            (0, 1),
            (0, 2),
            (0, 3),
            (1, 0),
            (1, 2),
            (1, 3),
            (2, 0),
            (2, 1),
            (2, 2),
            (2, 3),
            (3, 1),
        ]);
        let range = Rect::at(0, 1, 3, 3).expect("bounded test range");

        let addresses = store
            .cells(range)
            .map(|(address, _cell)| address)
            .collect::<Vec<_>>();

        assert_eq!(
            addresses,
            vec![
                Address::at(0, 1).unwrap(),
                Address::at(0, 2).unwrap(),
                Address::at(1, 2).unwrap(),
                Address::at(2, 1).unwrap(),
                Address::at(2, 2).unwrap(),
            ]
        );
    }

    #[test]
    fn range_cells_handle_sparse_rows_and_grid_edges() {
        let store = empty_store(&[(0, 0), (2, 0), (2, 16_383), (3, 0), (1_048_575, 16_383)]);

        let middle = Rect::at(1, 0, 4, 1).expect("bounded middle range");
        assert_eq!(
            store
                .cells(middle)
                .map(|(address, _cell)| address)
                .collect::<Vec<_>>(),
            vec![Address::at(2, 0).unwrap(), Address::at(3, 0).unwrap()]
        );

        let last = Rect::single(Address::at(1_048_575, 16_383).unwrap());
        assert_eq!(
            store
                .cells(last)
                .map(|(address, _cell)| address)
                .collect::<Vec<_>>(),
            vec![Address::at(1_048_575, 16_383).unwrap()]
        );
    }

    #[test]
    fn numbers_preserve_lexemes_and_convert_explicitly() {
        let number = Number::new("  -0.000  ").expect("valid number");
        assert_eq!(number.as_str(), "  -0.000  ");
        assert_eq!(number.as_f64(), Some(-0.0));
        assert!(Number::new("NaN").is_err());
        assert!(Number::new("not a number").is_err());
        assert!(Number::try_from(f64::INFINITY).is_err());
    }

    #[test]
    fn edit_content_keeps_text_inert_and_formulas_explicit() {
        assert!(matches!(
            Content::from("=SUM(A1:A3)"),
            Content::Value(Value::Text(text)) if text.as_str() == "=SUM(A1:A3)"
        ));
        assert!(matches!(
            Content::from(42_i32),
            Content::Value(Value::Number(number)) if number.as_str() == "42"
        ));
        assert!(matches!(
            Content::from(Formula::new("SUM(A1:A3)").expect("formula")),
            Content::Formula(_)
        ));
    }

    #[test]
    fn dates_are_checked_and_keep_their_lexical_form() {
        let date = Date::new("2026-07-31T12:34:56.250-07:00").expect("date");
        assert_eq!(date.as_str(), "2026-07-31T12:34:56.250-07:00");
        assert!(Date::new("2026-02-29").is_err());
        assert!(Value::date("not a date").is_err());
    }

    #[test]
    fn producer_unknown_error_values_are_read_only() {
        let content = Content::Value(Value::Error(ErrorValue::Unknown("#VENDOR!".into())));
        assert!(content.validate_for_write().is_err());
    }
}
