//! Dependency-free Numbers table semantics.
//!
//! The native archive reader owns protobuf traversal and producer-specific
//! sidecars. This module owns the checked sparse value model. Coordinates are
//! compact `u32` pairs, builders use one contiguous vector, and finished
//! tables retain one immutable boxed slice. No dense grid is allocated by the
//! semantic model.

use crate::cell::Value;
use std::fmt;

/// A checked zero-based table coordinate.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Position {
    row: u32,
    column: u32,
}

impl Position {
    /// Creates a zero-based coordinate.
    #[must_use]
    pub const fn new(row: u32, column: u32) -> Self {
        Self { row, column }
    }

    /// Converts platform-sized coordinates without truncation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::CoordinateOverflow`] if either coordinate does not
    /// fit in the compact representation.
    pub fn try_from_usize(row: usize, column: usize) -> Result<Self> {
        let compact_row =
            u32::try_from(row).map_err(|_conversion| Error::CoordinateOverflow { row, column })?;
        let compact_column = u32::try_from(column)
            .map_err(|_conversion| Error::CoordinateOverflow { row, column })?;
        Ok(Self::new(compact_row, compact_column))
    }

    /// Returns the zero-based row.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }

    /// Returns the zero-based column.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }
}

/// The declared addressable extent of a table.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Dimensions {
    rows: u32,
    columns: u32,
}

impl Dimensions {
    /// Creates a table extent without allocating.
    #[must_use]
    pub const fn new(rows: u32, columns: u32) -> Self {
        Self { rows, columns }
    }

    /// Converts platform-sized dimensions without truncation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::CoordinateOverflow`] if either dimension does not fit
    /// in the compact representation.
    pub fn try_from_usize(rows: usize, columns: usize) -> Result<Self> {
        let rows_u32 = u32::try_from(rows).map_err(|_conversion| Error::CoordinateOverflow {
            row: rows,
            column: columns,
        })?;
        let columns_u32 =
            u32::try_from(columns).map_err(|_conversion| Error::CoordinateOverflow {
                row: rows,
                column: columns,
            })?;
        Ok(Self::new(rows_u32, columns_u32))
    }

    /// Returns the declared row count.
    #[must_use]
    pub const fn rows(self) -> u32 {
        self.rows
    }

    /// Returns the declared column count.
    #[must_use]
    pub const fn columns(self) -> u32 {
        self.columns
    }

    /// Returns the dense area when it fits in `usize`.
    #[must_use]
    pub const fn area(self) -> Option<usize> {
        (self.rows as usize).checked_mul(self.columns as usize)
    }

    /// Creates a checked zero-based, half-open rectangular range.
    ///
    /// # Errors
    ///
    /// Returns an error when the range is inverted or extends beyond this
    /// table's declared extent.
    pub fn range(self, start: Position, end: Position) -> Result<Range> {
        let range = Range::new(start, end)?;
        if end.row > self.rows || end.column > self.columns {
            return Err(Error::OutOfBounds {
                position: end,
                dimensions: self,
            });
        }
        Ok(range)
    }

    fn contains(self, position: Position) -> bool {
        position.row < self.rows && position.column < self.columns
    }
}

/// A zero-based, half-open rectangular table range.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Range {
    start: Position,
    end: Position,
}

impl Range {
    /// Creates a range without a table-bound check.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidRange`] if the end coordinate precedes the
    /// start coordinate on either axis.
    pub fn new(start: Position, end: Position) -> Result<Self> {
        if start.row > end.row || start.column > end.column {
            return Err(Error::InvalidRange { start, end });
        }
        Ok(Self { start, end })
    }

    /// Returns the inclusive start and exclusive end coordinates.
    #[must_use]
    pub const fn bounds(self) -> (Position, Position) {
        (self.start, self.end)
    }

    /// Returns the first coordinate.
    #[must_use]
    pub const fn start(self) -> Position {
        self.start
    }

    /// Returns the exclusive end coordinate.
    #[must_use]
    pub const fn end(self) -> Position {
        self.end
    }

    /// Returns whether the range contains no cells.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.start.row == self.end.row || self.start.column == self.end.column
    }

    /// Returns the number of cells when it fits in `usize`.
    #[must_use]
    pub const fn area(self) -> Option<usize> {
        let rows = (self.end.row - self.start.row) as usize;
        let columns = (self.end.column - self.start.column) as usize;
        rows.checked_mul(columns)
    }

    /// Returns whether a coordinate belongs to the range.
    #[must_use]
    pub const fn contains(self, position: Position) -> bool {
        position.row >= self.start.row
            && position.row < self.end.row
            && position.column >= self.start.column
            && position.column < self.end.column
    }
}

/// One materialized sparse cell.
#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    position: Position,
    value: Value,
}

impl Cell {
    /// Creates a cell record.
    #[must_use]
    pub const fn new(position: Position, value: Value) -> Self {
        Self { position, value }
    }

    /// Returns the cell coordinate.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Borrows the cell value.
    #[must_use]
    pub const fn value(&self) -> &Value {
        &self.value
    }

    /// Returns the owned value.
    #[must_use]
    pub fn into_value(self) -> Value {
        self.value
    }
}

/// The result of looking up a coordinate without conflating missing and
/// explicitly stored empty cells.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum View<'a> {
    /// No semantic cell is stored at this coordinate.
    Missing,
    /// A semantic cell is stored, including [`Value::Empty`].
    Stored(&'a Value),
    /// The coordinate is covered by a format-owned merged region.
    ///
    /// Merge coverage is supplied by a concrete format sidecar; the leaf
    /// model never materializes follower cells.
    Covered,
}

/// Errors returned by checked sparse table operations.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    /// A coordinate is outside the table's declared extent.
    OutOfBounds {
        /// Requested coordinate.
        position: Position,
        /// Declared table extent.
        dimensions: Dimensions,
    },
    /// A range has descending bounds.
    InvalidRange {
        /// Inclusive range start.
        start: Position,
        /// Exclusive range end.
        end: Position,
    },
    /// A coordinate or dimension cannot fit in the compact representation.
    CoordinateOverflow {
        /// Row coordinate or dimension.
        row: usize,
        /// Column coordinate or dimension.
        column: usize,
    },
    /// A dense view exceeds its caller-supplied materialization budget.
    BudgetExceeded {
        /// Requested cell count.
        requested: usize,
        /// Maximum permitted cell count.
        maximum: usize,
    },
    /// A builder received two cells at the same coordinate.
    DuplicatePosition {
        /// Repeated coordinate.
        position: Position,
    },
    /// A sheet received two tables with the same semantic name.
    DuplicateTableName {
        /// Repeated table name.
        name: String,
    },
    /// A vector reservation failed before mutation.
    Allocation {
        /// Collection being reserved.
        resource: &'static str,
        /// Requested additional elements.
        amount: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::OutOfBounds {
                position,
                dimensions,
            } => write!(
                formatter,
                "table coordinate ({}, {}) is outside {}x{}",
                position.row, position.column, dimensions.rows, dimensions.columns
            ),
            Self::InvalidRange { start, end } => write!(
                formatter,
                "table range ({}, {})..({}, {}) is inverted",
                start.row, start.column, end.row, end.column
            ),
            Self::CoordinateOverflow { row, column } => {
                write!(
                    formatter,
                    "table coordinate ({row}, {column}) overflows u32"
                )
            },
            Self::BudgetExceeded { requested, maximum } => write!(
                formatter,
                "dense table view requests {requested} cells, budget is {maximum}"
            ),
            Self::DuplicatePosition { position } => write!(
                formatter,
                "table contains duplicate coordinate ({}, {})",
                position.row, position.column
            ),
            Self::DuplicateTableName { name } => {
                write!(formatter, "sheet contains duplicate table name {name:?}")
            },
            Self::Allocation { resource, amount } => {
                write!(
                    formatter,
                    "table allocation failed for {resource}: {amount}"
                )
            },
        }
    }
}

impl std::error::Error for Error {}

/// Result type for checked table operations.
pub type Result<T> = std::result::Result<T, Error>;

/// An insertion failure that returns the rejected owned value to its caller.
#[derive(Debug, PartialEq)]
pub struct InsertError<T> {
    error: Error,
    value: T,
}

/// Result type for ownership-preserving insertion operations.
pub type InsertResult<T, V> = std::result::Result<T, InsertError<V>>;

impl<T> InsertError<T> {
    pub(crate) fn new(error: Error, value: T) -> Self {
        Self { error, value }
    }

    /// Borrows the insertion error.
    #[must_use]
    pub const fn error(&self) -> &Error {
        &self.error
    }

    /// Returns the error and rejected value.
    #[must_use]
    pub fn into_parts(self) -> (Error, T) {
        (self.error, self.value)
    }
}

/// A caller-supplied limit for dense table materialization.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct GridBudget {
    max_cells: usize,
}

impl GridBudget {
    /// Creates a dense-view budget in cells.
    #[must_use]
    pub const fn new(max_cells: usize) -> Self {
        Self { max_cells }
    }

    /// Returns the maximum permitted cell count.
    #[must_use]
    pub const fn max_cells(self) -> usize {
        self.max_cells
    }
}

/// A checked dense view over a bounded range.
#[derive(Debug, Clone, Copy)]
pub struct Grid<'a> {
    table: &'a Table,
    range: Range,
}

impl<'a> Grid<'a> {
    /// Returns the view's bounded range.
    #[must_use]
    pub const fn range(self) -> Range {
        self.range
    }

    /// Looks up one coordinate inside this view.
    #[must_use]
    pub fn view(self, position: Position) -> Option<View<'a>> {
        self.range
            .contains(position)
            .then(|| self.table.view(position))
    }

    /// Iterates row-major over every cell slot in the bounded range.
    #[must_use]
    pub fn iter(self) -> GridIter<'a> {
        GridIter {
            table: self.table,
            range: self.range,
            next: (!self.range.is_empty()).then_some(self.range.start),
        }
    }
}

/// Row-major iterator over a checked dense view.
#[derive(Debug, Clone)]
pub struct GridIter<'a> {
    table: &'a Table,
    range: Range,
    next: Option<Position>,
}

impl<'a> Iterator for GridIter<'a> {
    type Item = View<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        let position = self.next?;
        let column = position.column + 1;
        if column < self.range.end.column {
            self.next = Some(Position::new(position.row, column));
        } else {
            let row = position.row + 1;
            self.next =
                (row < self.range.end.row).then_some(Position::new(row, self.range.start.column));
        }
        Some(self.table.view(position))
    }
}

/// An immutable, sparse semantic table.
#[derive(Debug, Clone, PartialEq)]
pub struct Table {
    name: Box<str>,
    dimensions: Dimensions,
    cells: Box<[Cell]>,
    column_headers: Box<[String]>,
    row_headers: Box<[String]>,
}

impl Table {
    /// Creates an empty immutable table with a declared extent.
    #[must_use]
    pub fn new(name: impl Into<String>, dimensions: Dimensions) -> Self {
        Self {
            name: name.into().into_boxed_str(),
            dimensions,
            cells: Box::new([]),
            column_headers: Box::new([]),
            row_headers: Box::new([]),
        }
    }

    /// Creates a mutable builder for a table.
    #[must_use]
    pub fn builder(name: impl Into<String>, dimensions: Dimensions) -> Builder {
        Builder::new(name, dimensions)
    }

    /// Borrows the table name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the declared extent.
    #[must_use]
    pub const fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

    /// Returns the declared row count.
    #[must_use]
    pub const fn row_count(&self) -> u32 {
        self.dimensions.rows
    }

    /// Returns the declared column count.
    #[must_use]
    pub const fn column_count(&self) -> u32 {
        self.dimensions.columns
    }

    /// Borrows a materialized value at a coordinate.
    #[must_use]
    pub fn get(&self, position: Position) -> Option<&Value> {
        match self.view(position) {
            View::Stored(value) => Some(value),
            View::Missing | View::Covered => None,
        }
    }

    /// Returns the compact stored/missing view for a coordinate.
    #[must_use]
    pub fn view(&self, position: Position) -> View<'_> {
        self.cells
            .binary_search_by_key(&position, Cell::position)
            .map_or(View::Missing, |index| {
                View::Stored(self.cells[index].value())
            })
    }

    /// Iterates over sparse cells in row-major order within a range.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidRange`] or [`Error::OutOfBounds`] when the
    /// range is not contained by the table's declared extent.
    pub fn cells(&self, range: Range) -> Result<impl Iterator<Item = &Cell> + '_> {
        let checked_range = self.dimensions.range(range.start, range.end)?;
        let first = self
            .cells
            .partition_point(|cell| cell.position() < checked_range.start);
        Ok(self.cells[first..]
            .iter()
            .take_while(move |cell| checked_range.contains(cell.position())))
    }

    /// Returns the number of materialized cells.
    #[must_use]
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Returns the number of materialized non-empty values.
    #[must_use]
    pub fn non_empty_cell_count(&self) -> usize {
        self.cells
            .iter()
            .filter(|cell| !cell.value().is_empty())
            .count()
    }

    /// Iterates over column headers in native order.
    #[must_use]
    pub fn column_headers(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.column_headers.iter().map(String::as_str)
    }

    /// Iterates over row headers in native order.
    #[must_use]
    pub fn row_headers(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.row_headers.iter().map(String::as_str)
    }

    /// Creates a dense view only when it fits the caller's budget.
    ///
    /// # Errors
    ///
    /// Returns [`Error::BudgetExceeded`] when the range area is larger than
    /// the supplied budget.
    pub fn grid(&self, range: Range, budget: GridBudget) -> Result<Grid<'_>> {
        let requested = range.area().ok_or(Error::BudgetExceeded {
            requested: usize::MAX,
            maximum: budget.max_cells,
        })?;
        if requested > budget.max_cells {
            return Err(Error::BudgetExceeded {
                requested,
                maximum: budget.max_cells,
            });
        }
        Ok(Grid { table: self, range })
    }

    /// Consumes the table and returns its immutable sparse cells.
    #[must_use]
    pub fn into_cells(self) -> Box<[Cell]> {
        self.cells
    }
}

/// A fallible mutable builder for an immutable sparse table.
#[derive(Debug, Clone, PartialEq)]
pub struct Builder {
    name: String,
    dimensions: Dimensions,
    cells: Vec<Cell>,
    column_headers: Vec<String>,
    row_headers: Vec<String>,
    sorted: bool,
}

impl Builder {
    /// Creates an empty builder with a declared extent.
    #[must_use]
    pub fn new(name: impl Into<String>, dimensions: Dimensions) -> Self {
        Self {
            name: name.into(),
            dimensions,
            cells: Vec::new(),
            column_headers: Vec::new(),
            row_headers: Vec::new(),
            sorted: true,
        }
    }

    /// Borrows the table name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the declared extent.
    #[must_use]
    pub const fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

    /// Changes the declared extent without moving sparse cells.
    ///
    /// # Errors
    ///
    /// Returns [`Error::OutOfBounds`] if a materialized cell would fall
    /// outside the requested extent.
    pub fn resize(&mut self, dimensions: Dimensions) -> Result<()> {
        if let Some(cell) = self
            .cells
            .iter()
            .find(|cell| !dimensions.contains(cell.position()))
        {
            return Err(Error::OutOfBounds {
                position: cell.position(),
                dimensions,
            });
        }
        self.dimensions = dimensions;
        Ok(())
    }

    /// Borrows a materialized value at a coordinate.
    #[must_use]
    pub fn get(&self, position: Position) -> Option<&Value> {
        if self.sorted {
            match self.cells.binary_search_by_key(&position, Cell::position) {
                Ok(index) => Some(self.cells[index].value()),
                Err(_) => None,
            }
        } else {
            self.cells
                .iter()
                .find(|cell| cell.position() == position)
                .map(Cell::value)
        }
    }

    /// Iterates over the builder's current sparse cells.
    #[must_use]
    pub fn cells(&self) -> impl ExactSizeIterator<Item = &Cell> + '_ {
        self.cells.iter()
    }

    /// Returns the number of materialized cells.
    #[must_use]
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Returns the number of materialized non-empty values.
    #[must_use]
    pub fn non_empty_cell_count(&self) -> usize {
        self.cells
            .iter()
            .filter(|cell| !cell.value().is_empty())
            .count()
    }

    /// Replaces or inserts one checked sparse value.
    ///
    /// # Errors
    ///
    /// Returns the rejected value when the coordinate is out of bounds or a
    /// reservation fails.
    pub fn set(
        &mut self,
        position: Position,
        value: Value,
    ) -> std::result::Result<(), InsertError<Value>> {
        if !self.dimensions.contains(position) {
            return Err(InsertError::new(
                Error::OutOfBounds {
                    position,
                    dimensions: self.dimensions,
                },
                value,
            ));
        }

        if !self.sorted {
            self.cells.sort_unstable_by_key(Cell::position);
            self.sorted = true;
        }

        match self.cells.last().map(Cell::position) {
            Some(last) if position > last => {
                if let Err(_allocation) = self.cells.try_reserve(1) {
                    return Err(InsertError::new(
                        Error::Allocation {
                            resource: "table cells",
                            amount: 1,
                        },
                        value,
                    ));
                }
                self.cells.push(Cell::new(position, value));
            },
            Some(last) if position == last => {
                if let Some(cell) = self.cells.last_mut() {
                    cell.value = value;
                }
            },
            _ => match self.cells.binary_search_by_key(&position, Cell::position) {
                Ok(index) => self.cells[index].value = value,
                Err(index) => {
                    if let Err(_allocation) = self.cells.try_reserve(1) {
                        return Err(InsertError::new(
                            Error::Allocation {
                                resource: "table cells",
                                amount: 1,
                            },
                            value,
                        ));
                    }
                    self.cells.insert(index, Cell::new(position, value));
                },
            },
        }
        Ok(())
    }

    /// Appends a cell for high-throughput archive ingestion.
    ///
    /// Appended cells are sorted once by [`Self::finish`]. Duplicate
    /// coordinates are rejected by the finish operation.
    ///
    /// # Errors
    ///
    /// Returns the rejected cell when its coordinate is outside the declared
    /// extent or a reservation fails.
    pub fn push(&mut self, cell: Cell) -> std::result::Result<(), InsertError<Cell>> {
        if !self.dimensions.contains(cell.position()) {
            return Err(InsertError::new(
                Error::OutOfBounds {
                    position: cell.position(),
                    dimensions: self.dimensions,
                },
                cell,
            ));
        }
        if let Some(last) = self.cells.last().map(Cell::position) {
            self.sorted &= cell.position() >= last;
        }
        if let Err(_allocation) = self.cells.try_reserve(1) {
            return Err(InsertError::new(
                Error::Allocation {
                    resource: "table cells",
                    amount: 1,
                },
                cell,
            ));
        }
        self.cells.push(cell);
        Ok(())
    }

    /// Replaces column headers.
    ///
    /// # Errors
    ///
    /// Returns an allocation error before replacing the current headers.
    pub fn set_column_headers<I, S>(&mut self, headers: I) -> Result<()>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        let header_iter = headers.into_iter();
        let (lower_bound, _) = header_iter.size_hint();
        let mut values = Vec::new();
        values
            .try_reserve(lower_bound)
            .map_err(|_allocation| Error::Allocation {
                resource: "table column headers",
                amount: lower_bound,
            })?;
        values.extend(header_iter.map(Into::into));
        self.column_headers = values;
        Ok(())
    }

    /// Iterates over column headers in native order.
    #[must_use]
    pub fn column_headers(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.column_headers.iter().map(String::as_str)
    }

    /// Replaces row headers.
    ///
    /// # Errors
    ///
    /// Returns an allocation error before replacing the current headers.
    pub fn set_row_headers<I, S>(&mut self, headers: I) -> Result<()>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        let header_iter = headers.into_iter();
        let (lower_bound, _) = header_iter.size_hint();
        let mut values = Vec::new();
        values
            .try_reserve(lower_bound)
            .map_err(|_allocation| Error::Allocation {
                resource: "table row headers",
                amount: lower_bound,
            })?;
        values.extend(header_iter.map(Into::into));
        self.row_headers = values;
        Ok(())
    }

    /// Iterates over row headers in native order.
    #[must_use]
    pub fn row_headers(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.row_headers.iter().map(String::as_str)
    }

    /// Consumes the builder and returns its sorted sparse cells.
    #[must_use]
    pub fn into_cells(mut self) -> Box<[Cell]> {
        self.cells.sort_unstable_by_key(Cell::position);
        self.cells.into_boxed_slice()
    }

    /// Sorts and seals the builder into an immutable table.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DuplicatePosition`] if cells appended through
    /// [`Self::push`] contain the same coordinate more than once.
    pub fn finish(mut self) -> Result<Table> {
        self.cells.sort_unstable_by_key(Cell::position);
        for pair in self.cells.windows(2) {
            if pair[0].position() == pair[1].position() {
                return Err(Error::DuplicatePosition {
                    position: pair[0].position(),
                });
            }
        }
        Ok(Table {
            name: self.name.into_boxed_str(),
            dimensions: self.dimensions,
            cells: self.cells.into_boxed_slice(),
            column_headers: self.column_headers.into_boxed_slice(),
            row_headers: self.row_headers.into_boxed_slice(),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn compact_sparse_cells_are_sorted_and_replaced_without_duplicates() {
        let mut table = Builder::new("Test", Dimensions::new(4, 4));
        assert!(table.set(Position::new(2, 1), Value::Number(2.0)).is_ok());
        assert!(table.set(Position::new(0, 3), Value::Number(1.0)).is_ok());
        assert!(table.set(Position::new(2, 1), Value::Number(3.0)).is_ok());

        let table = table.finish();
        assert!(table.is_ok());
        let table = table.unwrap_or_else(|error| panic!("unexpected table error: {error}"));
        let range = Range::new(Position::new(0, 0), Position::new(4, 4))
            .unwrap_or_else(|error| panic!("unexpected range error: {error}"));
        let positions: Vec<_> = table
            .cells(range)
            .unwrap_or_else(|error| panic!("unexpected cells error: {error}"))
            .map(Cell::position)
            .collect();
        assert_eq!(positions, [Position::new(0, 3), Position::new(2, 1)]);
        assert_eq!(table.get(Position::new(2, 1)), Some(&Value::Number(3.0)));
        assert_eq!(table.cell_count(), 2);
    }

    #[test]
    fn push_sorts_once_and_rejects_duplicate_coordinates() {
        let mut builder = Builder::new("Test", Dimensions::new(2, 2));
        assert!(
            builder
                .push(Cell::new(Position::new(1, 1), Value::Number(2.0)))
                .is_ok()
        );
        assert!(
            builder
                .push(Cell::new(Position::new(0, 0), Value::Number(1.0)))
                .is_ok()
        );
        let table = builder.finish();
        assert!(table.is_ok());

        let mut duplicate = Builder::new("Test", Dimensions::new(2, 2));
        assert!(
            duplicate
                .push(Cell::new(Position::new(0, 0), Value::Empty))
                .is_ok()
        );
        assert!(
            duplicate
                .push(Cell::new(Position::new(0, 0), Value::Number(1.0)))
                .is_ok()
        );
        assert!(matches!(
            duplicate.finish(),
            Err(Error::DuplicatePosition { .. })
        ));
    }

    #[test]
    fn views_distinguish_missing_and_explicit_empty() {
        let mut builder = Builder::new("Test", Dimensions::new(1, 2));
        assert!(builder.set(Position::new(0, 0), Value::Empty).is_ok());
        let table = builder
            .finish()
            .unwrap_or_else(|error| panic!("unexpected table error: {error}"));
        assert!(matches!(
            table.view(Position::new(0, 0)),
            View::Stored(&Value::Empty)
        ));
        assert!(matches!(table.view(Position::new(0, 1)), View::Missing));
    }

    #[test]
    fn bounded_grid_is_lazy_and_budgeted() {
        let table = Table::new("Test", Dimensions::new(2, 3));
        let range = Range::new(Position::new(0, 0), Position::new(2, 2))
            .unwrap_or_else(|error| panic!("unexpected range error: {error}"));
        assert!(table.grid(range, GridBudget::new(3)).is_err());
        let grid = table
            .grid(range, GridBudget::new(4))
            .unwrap_or_else(|error| panic!("unexpected grid error: {error}"));
        assert_eq!(grid.iter().count(), 4);
    }

    #[test]
    fn coordinates_and_ranges_are_checked() {
        let dimensions = Dimensions::new(2, 2);
        assert!(matches!(
            dimensions.range(Position::new(0, 0), Position::new(3, 2)),
            Err(Error::OutOfBounds { .. })
        ));
        assert!(matches!(
            Range::new(Position::new(1, 0), Position::new(0, 1)),
            Err(Error::InvalidRange { .. })
        ));
        assert!(matches!(
            Position::try_from_usize(usize::MAX, 0),
            Err(Error::CoordinateOverflow { .. })
        ));
    }
}
