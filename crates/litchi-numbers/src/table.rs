//! Numbers table compatibility facade.
//!
//! The sparse table model is owned by `litchi-iwa-common`. This module keeps
//! the established Numbers paths and adds the Numbers-only name selector while
//! forwarding all neutral table operations to the shared implementation.

/// Lossless table appearance values and selector-first transactions.
pub mod appearance;
/// Presence-preserving semantic cell reads.
pub mod cells;
/// Compact, archive-free cell coordinates and A1 selectors.
pub mod coordinate;
/// Checked row, column, and point-size values.
pub mod dimension;
/// Checked, bounded plans for applying multiple cell mutations.
pub mod edit;
/// Header, footer, and repeating-row/column semantics.
pub mod headers;
/// Archive-free interactive table-lock semantics.
pub mod lock;
/// Compact merged-cell geometry and topology algebra.
pub mod merge;
/// Exact-source transactions for moving a rooted table between sheets.
pub mod relocation;
/// Checked, archive-free table sort semantics.
pub mod sort;
/// Compact, presence-preserving table title semantics.
pub mod title;
/// Section-relative table topology edits.
pub mod topology;

use crate::cell::Value;
use crate::selector::TableSelector;

pub use litchi_iwa_common::table::coordinate::{
    AddressError, CellPosition, CellRange, Error as CoordinateError,
};
pub use litchi_iwa_common::table::model::{
    Cell, Dimensions, Error, Grid, GridBudget, GridIter, InsertError, InsertResult, Result, View,
};

/// Narrow migration name for [`CellPosition`] used by existing archive
/// adapters. New semantic code should use the focused name.
pub type Position = CellPosition;

/// Narrow migration name for [`CellRange`] used by existing archive adapters.
/// New semantic code should use the focused name.
pub type Range = CellRange;

/// The format-independent table representation used by Numbers.
///
/// Numbers keeps this small facade so existing callers retain the exact-name
/// selector API. The owned sparse storage, bounded grid, and CSV projection
/// live in the shared common model.
#[derive(Debug, Clone, PartialEq)]
pub struct Table(litchi_iwa_common::table::model::Table);

/// The shared table representation accepted by host and cross-format APIs.
pub type SharedTable = litchi_iwa_common::table::model::Table;

impl Table {
    /// Creates an empty immutable table with a declared extent.
    #[must_use]
    pub fn new(name: impl Into<String>, dimensions: Dimensions) -> Self {
        Self(SharedTable::new(name, dimensions))
    }

    /// Creates a mutable builder for a table.
    #[must_use]
    pub fn builder(name: impl Into<String>, dimensions: Dimensions) -> Builder {
        Builder::new(name, dimensions)
    }

    /// Wraps the shared table model in the Numbers compatibility facade.
    #[must_use]
    pub fn from_shared(table: SharedTable) -> Self {
        Self(table)
    }

    /// Consumes the Numbers facade and returns the shared table model.
    #[must_use]
    pub fn into_shared(self) -> SharedTable {
        self.0
    }

    /// Borrows the underlying shared table model.
    #[must_use]
    pub const fn as_shared(&self) -> &SharedTable {
        &self.0
    }

    /// Borrows the table name.
    #[must_use]
    pub fn name(&self) -> &str {
        self.0.name()
    }

    /// Returns an exact-name selector for this table.
    ///
    /// Table names are matched case-sensitively within their owning sheet.
    /// The selector intentionally contains only the borrowed semantic name;
    /// native table or archive identifiers are not part of the value model.
    #[must_use]
    pub fn selector(&self) -> TableSelector<'_> {
        TableSelector::name(self.name())
    }

    /// Returns the declared extent.
    #[must_use]
    pub const fn dimensions(&self) -> Dimensions {
        self.0.dimensions()
    }

    /// Returns the declared row count.
    #[must_use]
    pub const fn row_count(&self) -> u32 {
        self.0.row_count()
    }

    /// Returns the declared column count.
    #[must_use]
    pub const fn column_count(&self) -> u32 {
        self.0.column_count()
    }

    /// Borrows a materialized value at a coordinate.
    #[must_use]
    pub fn get(&self, position: Position) -> Option<&Value> {
        self.0.get(position)
    }

    /// Looks up a materialized value by a checked A1 selector.
    pub fn get_a1(&self, address: &str) -> Result<Option<&Value>> {
        self.0.get_a1(address)
    }

    /// Looks up the presence-preserving view for a checked A1 selector.
    pub fn view_a1(&self, address: &str) -> Result<View<'_>> {
        self.0.view_a1(address)
    }

    /// Returns the compact stored/missing view for a coordinate.
    #[must_use]
    pub fn view(&self, position: Position) -> View<'_> {
        self.0.view(position)
    }

    /// Iterates over sparse cells in row-major order within a range.
    pub fn cells(&self, range: Range) -> Result<impl Iterator<Item = &Cell> + '_> {
        self.0.cells(range)
    }

    /// Iterates over sparse cells selected by a checked A1 range.
    pub fn cells_a1(&self, address: &str) -> Result<impl Iterator<Item = &Cell> + '_> {
        self.0.cells_a1(address)
    }

    /// Iterates over all materialized sparse cells in row-major order.
    #[must_use]
    pub fn iter_cells(&self) -> impl ExactSizeIterator<Item = &Cell> + '_ {
        self.0.iter_cells()
    }

    /// Returns the number of materialized cells.
    #[must_use]
    pub fn cell_count(&self) -> usize {
        self.0.cell_count()
    }

    /// Returns the number of materialized non-empty values.
    #[must_use]
    pub fn non_empty_cell_count(&self) -> usize {
        self.0.non_empty_cell_count()
    }

    /// Projects the sparse table to RFC 4180-compatible CSV text.
    #[must_use]
    pub fn to_csv(&self) -> String {
        self.0.to_csv()
    }

    /// Iterates over column headers in native order.
    #[must_use]
    pub fn column_headers(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.0.column_headers()
    }

    /// Iterates over row headers in native order.
    #[must_use]
    pub fn row_headers(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.0.row_headers()
    }

    /// Creates a dense view only when it fits the caller's budget.
    pub fn grid(&self, range: Range, budget: GridBudget) -> Result<Grid<'_>> {
        self.0.grid(range, budget)
    }

    /// Consumes the table and returns its immutable sparse cells.
    #[must_use]
    pub fn into_cells(self) -> Box<[Cell]> {
        self.0.into_cells()
    }
}

impl From<SharedTable> for Table {
    fn from(table: SharedTable) -> Self {
        Self::from_shared(table)
    }
}

impl From<Table> for SharedTable {
    fn from(table: Table) -> Self {
        table.into_shared()
    }
}

/// A fallible mutable builder for an immutable sparse table.
#[derive(Debug, Clone, PartialEq)]
pub struct Builder(litchi_iwa_common::table::model::Builder);

impl Builder {
    /// Creates an empty builder with a declared extent.
    #[must_use]
    pub fn new(name: impl Into<String>, dimensions: Dimensions) -> Self {
        Self(SharedTable::builder(name, dimensions))
    }

    /// Wraps a shared mutable table builder.
    #[must_use]
    pub fn from_shared(builder: litchi_iwa_common::table::model::Builder) -> Self {
        Self(builder)
    }

    /// Consumes the Numbers builder and returns the shared mutable builder.
    #[must_use]
    pub fn into_shared(self) -> litchi_iwa_common::table::model::Builder {
        self.0
    }

    /// Borrows the builder name.
    #[must_use]
    pub fn name(&self) -> &str {
        self.0.name()
    }

    /// Returns the declared extent.
    #[must_use]
    pub const fn dimensions(&self) -> Dimensions {
        self.0.dimensions()
    }

    /// Changes the declared extent without moving sparse cells.
    pub fn resize(&mut self, dimensions: Dimensions) -> Result<()> {
        self.0.resize(dimensions)
    }

    /// Borrows a materialized value at a coordinate.
    #[must_use]
    pub fn get(&self, position: Position) -> Option<&Value> {
        self.0.get(position)
    }

    /// Iterates over the builder's current sparse cells.
    #[must_use]
    pub fn cells(&self) -> impl ExactSizeIterator<Item = &Cell> + '_ {
        self.0.cells()
    }

    /// Returns the number of materialized cells.
    #[must_use]
    pub fn cell_count(&self) -> usize {
        self.0.cell_count()
    }

    /// Returns the number of materialized non-empty values.
    #[must_use]
    pub fn non_empty_cell_count(&self) -> usize {
        self.0.non_empty_cell_count()
    }

    /// Replaces or inserts one checked sparse value.
    pub fn set(
        &mut self,
        position: Position,
        value: Value,
    ) -> std::result::Result<(), InsertError<Value>> {
        self.0.set(position, value)
    }

    /// Replaces or inserts one value selected by a checked A1 address.
    pub fn set_a1(
        &mut self,
        address: &str,
        value: Value,
    ) -> std::result::Result<(), InsertError<Value>> {
        self.0.set_a1(address, value)
    }

    /// Appends a cell for high-throughput archive ingestion.
    pub fn push(&mut self, cell: Cell) -> std::result::Result<(), InsertError<Cell>> {
        self.0.push(cell)
    }

    /// Replaces column headers.
    pub fn set_column_headers<I, S>(&mut self, headers: I) -> Result<()>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.0.set_column_headers(headers)
    }

    /// Iterates over column headers in native order.
    #[must_use]
    pub fn column_headers(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.0.column_headers()
    }

    /// Replaces row headers.
    pub fn set_row_headers<I, S>(&mut self, headers: I) -> Result<()>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.0.set_row_headers(headers)
    }

    /// Iterates over row headers in native order.
    #[must_use]
    pub fn row_headers(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.0.row_headers()
    }

    /// Consumes the builder and returns its sorted sparse cells.
    #[must_use]
    pub fn into_cells(self) -> Box<[Cell]> {
        self.0.into_cells()
    }

    /// Sorts and seals the builder into an immutable table.
    pub fn finish(self) -> Result<Table> {
        self.0.finish().map(Table::from_shared)
    }
}

impl From<litchi_iwa_common::table::model::Builder> for Builder {
    fn from(builder: litchi_iwa_common::table::model::Builder) -> Self {
        Self::from_shared(builder)
    }
}

impl From<Builder> for litchi_iwa_common::table::model::Builder {
    fn from(builder: Builder) -> Self {
        builder.into_shared()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn table_selector_is_an_exact_case_sensitive_semantic_name() {
        let table = Table::new("Revenue", Dimensions::new(1, 1));

        assert_eq!(table.selector(), TableSelector::name("Revenue"));
        assert_ne!(table.selector(), TableSelector::name("revenue"));
    }

    #[test]
    fn shared_conversion_preserves_table_contents() {
        let mut builder = Table::builder("Revenue", Dimensions::new(2, 2));
        assert!(
            builder
                .set_a1("B2", Value::Text("value".to_owned()))
                .is_ok()
        );
        let table = builder
            .finish()
            .unwrap_or_else(|error| panic!("unexpected table error: {error}"));
        let shared = table.clone().into_shared();
        let restored = Table::from_shared(shared.clone());
        assert_eq!(restored.name(), "Revenue");
        assert_eq!(
            restored.get_a1("B2"),
            Ok(Some(&Value::Text("value".to_owned())))
        );
        assert_eq!(restored.as_shared(), &shared);
    }
}
