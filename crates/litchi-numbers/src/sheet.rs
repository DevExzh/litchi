//! Dependency-free Numbers sheet semantics.

use super::table::{Error, InsertError, InsertResult, Table};

/// An immutable Numbers sheet containing semantic tables.
#[derive(Debug, Clone, PartialEq)]
pub struct Sheet {
    name: Box<str>,
    index: usize,
    tables: Box<[Table]>,
}

impl Sheet {
    /// Creates an empty immutable sheet.
    #[must_use]
    pub fn new(name: impl Into<String>, index: usize) -> Self {
        Self {
            name: name.into().into_boxed_str(),
            index,
            tables: Box::new([]),
        }
    }

    /// Creates a mutable builder for a sheet.
    #[must_use]
    pub fn builder(name: impl Into<String>, index: usize) -> Builder {
        Builder::new(name, index)
    }

    /// Borrows the sheet name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the zero-based sheet index.
    #[must_use]
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Iterates over tables in native order.
    #[must_use]
    pub fn tables(&self) -> impl ExactSizeIterator<Item = &Table> + '_ {
        self.tables.iter()
    }

    /// Returns a table by name.
    #[must_use]
    pub fn table(&self, name: &str) -> Option<&Table> {
        self.tables.iter().find(|table| table.name() == name)
    }

    /// Returns whether the sheet has no tables.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.tables.is_empty()
    }

    /// Returns the number of tables.
    #[must_use]
    pub fn table_count(&self) -> usize {
        self.tables.len()
    }

    /// Returns the dense addressable cell count when it fits in `usize`.
    #[must_use]
    pub fn addressable_cell_count(&self) -> Option<usize> {
        self.tables.iter().try_fold(0usize, |total, table| {
            total.checked_add(table.dimensions().area()?)
        })
    }

    /// Returns the total number of materialized cells across all tables.
    #[must_use]
    pub fn materialized_cell_count(&self) -> usize {
        self.tables.iter().map(Table::cell_count).sum()
    }
}

/// A fallible mutable builder for an immutable sheet.
#[derive(Debug, Clone, PartialEq)]
pub struct Builder {
    name: String,
    index: usize,
    tables: Vec<Table>,
}

impl Builder {
    /// Creates an empty sheet builder.
    #[must_use]
    pub fn new(name: impl Into<String>, index: usize) -> Self {
        Self {
            name: name.into(),
            index,
            tables: Vec::new(),
        }
    }

    /// Borrows the sheet name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the zero-based sheet index.
    #[must_use]
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Adds a table without losing it on allocation failure.
    ///
    /// # Errors
    ///
    /// Returns the rejected table if its name duplicates an existing table or
    /// the backing vector cannot reserve one more entry.
    pub fn push_table(&mut self, table: Table) -> InsertResult<(), Table> {
        if self
            .tables
            .iter()
            .any(|stored| stored.name() == table.name())
        {
            return Err(InsertError::new(
                Error::DuplicateTableName {
                    name: table.name().to_owned(),
                },
                table,
            ));
        }
        if let Err(_allocation) = self.tables.try_reserve(1) {
            return Err(InsertError::new(
                Error::Allocation {
                    resource: "sheet tables",
                    amount: 1,
                },
                table,
            ));
        }
        self.tables.push(table);
        Ok(())
    }

    /// Seals the builder into an immutable sheet.
    #[must_use]
    pub fn finish(self) -> Sheet {
        Sheet {
            name: self.name.into_boxed_str(),
            index: self.index,
            tables: self.tables.into_boxed_slice(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::Value;
    use crate::table::{Builder as TableBuilder, Dimensions, Position};

    #[test]
    fn sheets_store_immutable_tables_without_peer_dependencies() {
        let mut table_builder = TableBuilder::new("Table 1", Dimensions::new(2, 3));
        assert!(
            table_builder
                .set(Position::new(1, 1), Value::Number(42.0))
                .is_ok()
        );
        let table = table_builder
            .finish()
            .unwrap_or_else(|error| panic!("unexpected table error: {error}"));

        let mut sheet_builder = Builder::new("Sheet 1", 0);
        assert!(sheet_builder.push_table(table).is_ok());
        let sheet = sheet_builder.finish();
        assert_eq!(sheet.name(), "Sheet 1");
        assert_eq!(sheet.table_count(), 1);
        assert_eq!(sheet.addressable_cell_count(), Some(6));
        assert_eq!(sheet.materialized_cell_count(), 1);
        assert!(sheet.table("Table 1").is_some());
    }
}
