//! Table expansion utilities for handling repeated cells and rows.
//!
//! ODF files can have cells and rows marked as "repeated" for efficiency.
//! This module provides utilities to expand these repeated elements into
//! their full representation.

use super::table::{Table, TableCell, TableRow};
use litchi_core::{Error, Result};

const MAX_EXPANDED_ITEMS: usize = 1_000_000;

/// Utilities for expanding repeated table elements
pub struct TableExpander;

impl TableExpander {
    /// Expand a table by resolving all repeated cells and rows.
    ///
    /// In ODF files, cells and rows can have `table:number-columns-repeated` and
    /// `table:number-rows-repeated` attributes to indicate repetition. This method
    /// expands all such repetitions into explicit cells and rows.
    ///
    /// # Arguments
    ///
    /// * `table` - The table to expand
    ///
    /// # Returns
    ///
    /// A new `Table` with all repetitions expanded
    ///
    /// # Example
    ///
    /// ```no_run
    /// use litchi_odt::elements::table_expansion::TableExpander;
    /// use litchi_odt::elements::table::Table;
    ///
    /// let table = Table::new();
    /// // ... populate table with repeated cells/rows ...
    ///
    /// let expanded = TableExpander::expand_table(&table).unwrap();
    /// ```
    pub fn expand_table(table: &Table) -> Result<Table> {
        let mut expanded_table = Table::try_new()?;

        // Copy table metadata
        if let Some(name) = table.name() {
            expanded_table.try_set_name(name)?;
        }
        if let Some(style) = table.style_name() {
            expanded_table.try_set_style_name(style)?;
        }

        // Expand all rows
        let mut expanded_rows = 0usize;
        for row in table.rows()? {
            // Check for row repetition
            let row_repeat_count = Self::get_row_repeat_count(&row)?;
            expanded_rows = expanded_rows.checked_add(row_repeat_count).ok_or_else(|| {
                Error::InvalidFormat("ODT expanded table row count overflow".to_string())
            })?;
            if expanded_rows > MAX_EXPANDED_ITEMS {
                return Err(Error::InvalidFormat(
                    "ODT expanded table exceeds 1000000 rows".to_string(),
                ));
            }

            for _ in 0..row_repeat_count {
                let expanded_row = Self::expand_row(&row)?;
                expanded_table.try_add_row(expanded_row)?;
            }
        }

        Ok(expanded_table)
    }

    /// Expand a single row by resolving all repeated cells.
    ///
    /// # Arguments
    ///
    /// * `row` - The row to expand
    ///
    /// # Returns
    ///
    /// A new `TableRow` with all repeated cells expanded
    fn expand_row(row: &TableRow) -> Result<TableRow> {
        let mut expanded_row = TableRow::try_new()?;

        // Copy row metadata
        if let Some(style) = row.style_name() {
            expanded_row.try_set_style_name(style)?;
        }

        // Expand all cells
        let mut expanded_cells = 0usize;
        for cell in row.cells()? {
            let repeat_count = Self::get_cell_repeat_count(&cell)?;
            expanded_cells = expanded_cells.checked_add(repeat_count).ok_or_else(|| {
                Error::InvalidFormat("ODT expanded table cell count overflow".to_string())
            })?;
            if expanded_cells > MAX_EXPANDED_ITEMS {
                return Err(Error::InvalidFormat(
                    "ODT expanded table row exceeds 1000000 cells".to_string(),
                ));
            }

            for _ in 0..repeat_count {
                let expanded_cell = Self::clone_cell(&cell)?;
                expanded_row.try_add_cell(expanded_cell)?;
            }
        }

        Ok(expanded_row)
    }

    /// Get the number of times a row should be repeated.
    ///
    /// Returns 1 if no repetition is specified.
    fn get_row_repeat_count(row: &TableRow) -> Result<usize> {
        let count = row.try_repeat_count()?;
        if count == 0 || count > MAX_EXPANDED_ITEMS {
            return Err(Error::InvalidFormat(
                "ODT table row repetition must be between 1 and 1000000".to_string(),
            ));
        }
        Ok(count)
    }

    /// Get the number of times a cell should be repeated.
    ///
    /// Returns 1 if no repetition is specified.
    fn get_cell_repeat_count(cell: &TableCell) -> Result<usize> {
        let count = cell.try_repeat_count()?;
        if count == 0 || count > MAX_EXPANDED_ITEMS {
            return Err(Error::InvalidFormat(
                "ODT table cell repetition must be between 1 and 1000000".to_string(),
            ));
        }
        Ok(count)
    }

    /// Clone a cell with all its properties.
    ///
    /// # Arguments
    ///
    /// * `cell` - The cell to clone
    ///
    /// # Returns
    ///
    /// A new `TableCell` with the same properties
    fn clone_cell(cell: &TableCell) -> Result<TableCell> {
        let mut new_cell = TableCell::try_new()?;

        // Copy text content
        new_cell.try_set_text(&cell.text()?)?;

        // Copy formula
        if let Some(formula) = cell.formula() {
            new_cell.try_set_formula(formula)?;
        }

        // Copy style
        if let Some(style) = cell.style_name() {
            new_cell.try_set_style_name(style)?;
        }

        // Copy span attributes (but not if they're from repetition)
        if cell.colspan() > 1 {
            new_cell.try_set_colspan(cell.colspan())?;
        }
        if cell.rowspan() > 1 {
            new_cell.try_set_rowspan(cell.rowspan())?;
        }

        Ok(new_cell)
    }

    /// Expand tables in-place, modifying the vector.
    ///
    /// This is a convenience method for expanding multiple tables.
    ///
    /// # Arguments
    ///
    /// * `tables` - Vector of tables to expand
    ///
    /// # Returns
    ///
    /// A new vector of expanded tables
    pub fn expand_tables(tables: Vec<Table>) -> Result<Vec<Table>> {
        let mut expanded_tables = Vec::new();
        expanded_tables
            .try_reserve_exact(tables.len())
            .map_err(|source| Error::Allocation {
                resource: "ODT expanded tables",
                source,
            })?;

        for table in tables {
            expanded_tables.push(Self::expand_table(&table)?);
        }

        Ok(expanded_tables)
    }
}

/// Extension trait for Table to add expansion methods
#[allow(dead_code)] // Library API extension trait
pub trait TableExpansionExt {
    /// Expand this table by resolving all repeated cells and rows.
    fn expand(&self) -> Result<Table>;
}

impl TableExpansionExt for Table {
    fn expand(&self) -> Result<Table> {
        TableExpander::expand_table(self)
    }
}
