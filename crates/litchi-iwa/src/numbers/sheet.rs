//! Archive-bound Numbers sheet conversion.
//!
//! The semantic sheet owner lives in `litchi_numbers::sheet`. This module is
//! deliberately crate-private: it retains only the temporary mutable state
//! needed while decoding native archives and transfers finished tables to the
//! immutable leaf model at the boundary.

use super::table::{NumbersTable, map_table_error};

/// Archive adapter used while decoding one native sheet.
///
/// This type is intentionally opaque outside `litchi-iwa`; callers that need
/// archive-free semantics should use `NumbersDocument::semantic_sheets`.
#[derive(Debug, Clone)]
pub struct NumbersSheet {
    pub(crate) name: String,
    pub(crate) index: usize,
    pub(crate) tables: Vec<NumbersTable>,
}

impl NumbersSheet {
    /// Creates a native sheet adapter.
    pub(crate) fn new(name: String, index: usize) -> Self {
        Self {
            name,
            index,
            tables: Vec::new(),
        }
    }

    /// Adds an archive-decoded table to the adapter.
    pub(crate) fn add_table(&mut self, table: NumbersTable) {
        self.tables.push(table);
    }

    /// Borrows the native sheet name without exposing archive state.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the logical zero-based sheet position.
    #[must_use]
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Iterates native tables without exposing the backing collection.
    #[must_use]
    pub fn tables(&self) -> impl ExactSizeIterator<Item = &NumbersTable> + '_ {
        self.tables.iter()
    }

    /// Returns a native table by exact name.
    #[must_use]
    pub fn table(&self, name: &str) -> Option<&NumbersTable> {
        self.tables.iter().find(|table| table.name() == name)
    }

    /// Returns the number of native tables.
    pub(crate) fn table_count(&self) -> usize {
        self.tables.len()
    }

    /// Returns the number of addressable cells across all native tables.
    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) fn total_cell_count(&self) -> usize {
        self.tables
            .iter()
            .map(|table| table.row_count() * table.column_count())
            .sum()
    }

    /// Consumes the archive adapter into the canonical dependency-free sheet.
    pub(crate) fn into_semantic(self) -> crate::Result<litchi_numbers::Sheet> {
        let mut builder = litchi_numbers::sheet::Builder::new(self.name, self.index);
        for table in self.tables {
            let semantic = table.into_semantic()?;
            builder
                .push_table(semantic)
                .map_err(|failure| map_table_error(failure.into_parts().0))?;
        }
        Ok(builder.finish())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::cell::CellValue;

    fn cell_number(value: f64) -> CellValue {
        CellValue::number(value).expect("finite test number")
    }

    #[test]
    fn adapter_is_not_a_public_semantic_model() {
        let sheet = NumbersSheet::new("Sheet1".to_owned(), 0);
        assert_eq!(sheet.name, "Sheet1");
        assert_eq!(sheet.index, 0);
        assert!(sheet.tables.is_empty());
        assert_eq!(sheet.total_cell_count(), 0);
    }

    #[test]
    fn adapter_collects_native_tables() {
        let mut sheet = NumbersSheet::new("Sheet1".to_owned(), 0);
        let mut table = NumbersTable::new("Table1");
        assert!(table.set_cell(0, 0, cell_number(1.0)).is_ok());

        sheet.add_table(table);

        assert_eq!(sheet.tables.len(), 1);
        assert_eq!(sheet.total_cell_count(), 1);
    }

    #[test]
    fn semantic_sheet_consumes_tables_without_rebuilding_cell_maps() {
        let mut sheet = NumbersSheet::new("Sheet1".to_owned(), 0);
        let mut table = NumbersTable::new("Table1");
        assert!(table.set_cell(1, 2, cell_number(42.0)).is_ok());
        assert!(table.set_column_headers(["A", "B", "C"]).is_ok());
        sheet.add_table(table);

        let semantic = sheet
            .into_semantic()
            .unwrap_or_else(|error| panic!("unexpected semantic conversion failure: {error}"));
        assert_eq!(semantic.name(), "Sheet1");
        assert_eq!(semantic.index(), 0);
        assert_eq!(semantic.table_count(), 1);
        let table = semantic
            .get("Table1")
            .unwrap_or_else(|error| panic!("unexpected selector failure: {error}"))
            .unwrap_or_else(|| panic!("converted table is missing"));
        assert_eq!(table.dimensions(), litchi_numbers::Dimensions::new(2, 3));
        assert_eq!(
            table.get(litchi_numbers::Position::new(1, 2)),
            Some(&cell_number(42.0))
        );
        assert_eq!(table.column_headers().collect::<Vec<_>>(), ["A", "B", "C"]);
    }
}
