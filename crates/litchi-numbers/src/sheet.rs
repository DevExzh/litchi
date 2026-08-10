//! Dependency-free Numbers sheet semantics.

use std::collections::HashSet;

use super::selector::TableSelector;
use super::table::{Error, InsertError, InsertResult, Table};

/// Errors raised while resolving a semantic table selector.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SelectorError {
    /// A malformed source exposed the same exact table name more than once.
    DuplicateTableName { name: Box<str> },
}

impl std::fmt::Display for SelectorError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::DuplicateTableName { name } => {
                write!(formatter, "sheet contains duplicate table name {name:?}")
            },
        }
    }
}

impl std::error::Error for SelectorError {}

/// Result type for checked semantic sheet selectors.
pub type Result<T> = std::result::Result<T, SelectorError>;

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

    /// Build one immutable sheet from tables in their native order.
    ///
    /// Unlike repeatedly calling [`Builder::push_table`], this validates all
    /// names through one fallibly allocated index. Native adapters use this
    /// bulk boundary after decoding a complete sheet, so large valid sheets
    /// do not repeatedly rescan their already-admitted table names.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DuplicateTableName`] when two tables have the same
    /// exact name, or [`Error::Allocation`] when the temporary name index
    /// cannot be reserved.
    pub fn try_from_tables(
        name: impl Into<String>,
        index: usize,
        tables: Vec<Table>,
    ) -> std::result::Result<Self, Error> {
        let mut names = HashSet::new();
        names
            .try_reserve(tables.len())
            .map_err(|_allocation| Error::Allocation {
                resource: "sheet table-name index",
                amount: tables.len(),
            })?;
        for table in &tables {
            if !names.insert(table.name()) {
                return Err(Error::DuplicateTableName {
                    name: table.name().to_owned(),
                });
            }
        }
        Ok(Self {
            name: name.into().into_boxed_str(),
            index,
            tables: tables.into_boxed_slice(),
        })
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

    /// Resolves a checked table selector.
    ///
    /// Name lookup is exact and source-order lookup never indexes directly.
    /// Missing tables and out-of-range positions are represented by `None`.
    ///
    /// # Errors
    ///
    /// Returns [`SelectorError::DuplicateTableName`] if a malformed semantic
    /// model contains the requested name more than once.
    pub fn select(&self, selector: TableSelector<'_>) -> Result<Option<&Table>> {
        match selector {
            TableSelector::Name(name) => {
                let mut matches = self.tables.iter().filter(|table| table.name() == name);
                let Some(table) = matches.next() else {
                    return Ok(None);
                };
                if matches.next().is_some() {
                    return Err(SelectorError::DuplicateTableName { name: name.into() });
                }
                Ok(Some(table))
            },
            TableSelector::Index(index) => Ok(self.tables.get(index)),
        }
    }

    /// Returns a table by exact name.
    ///
    /// # Errors
    ///
    /// Returns [`SelectorError::DuplicateTableName`] if a malformed semantic
    /// model contains the requested name more than once.
    pub fn get(&self, name: &str) -> Result<Option<&Table>> {
        self.select(TableSelector::Name(name))
    }

    /// Returns a table by checked zero-based source position.
    ///
    /// # Errors
    ///
    /// This selector currently cannot fail for a valid position; the
    /// `Result` keeps the selector boundary explicit for future validation.
    pub fn at(&self, index: usize) -> Result<Option<&Table>> {
        self.select(TableSelector::Index(index))
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
                .set(
                    Position::new(1, 1),
                    Value::number(42.0).expect("finite test number"),
                )
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
        assert!(
            sheet
                .get("Table 1")
                .unwrap_or_else(|error| panic!("unexpected selector failure: {error}"))
                .is_some()
        );
        assert!(
            sheet
                .at(0)
                .unwrap_or_else(|error| panic!("unexpected selector failure: {error}"))
                .is_some()
        );
        assert!(
            sheet
                .at(1)
                .unwrap_or_else(|error| panic!("unexpected selector failure: {error}"))
                .is_none()
        );
        assert!(
            sheet
                .get("table 1")
                .unwrap_or_else(|error| panic!("unexpected selector failure: {error}"))
                .is_none()
        );

        let second = TableBuilder::new("Second", Dimensions::new(1, 1))
            .finish()
            .unwrap_or_else(|error| panic!("unexpected table error: {error}"));
        let mut builder = Builder::new("Sheet 1", 0);
        let first = TableBuilder::new("First", Dimensions::new(1, 1))
            .finish()
            .unwrap_or_else(|error| panic!("unexpected table error: {error}"));
        assert!(builder.push_table(first).is_ok());
        assert!(builder.push_table(second).is_ok());
        let sheet = builder.finish();
        assert_eq!(
            sheet
                .select(TableSelector::index(1))
                .unwrap()
                .unwrap()
                .name(),
            "Second"
        );
        assert_eq!(
            sheet
                .select(TableSelector::name("First"))
                .unwrap()
                .unwrap()
                .name(),
            "First"
        );
    }

    #[test]
    fn bulk_construction_checks_names_once_and_preserves_order() {
        let first = TableBuilder::new("First", Dimensions::new(1, 1))
            .finish()
            .unwrap_or_else(|error| panic!("first table should be valid: {error}"));
        let second = TableBuilder::new("Second", Dimensions::new(1, 1))
            .finish()
            .unwrap_or_else(|error| panic!("second table should be valid: {error}"));
        let sheet = Sheet::try_from_tables("Sheet", 3, vec![first, second])
            .unwrap_or_else(|error| panic!("bulk sheet should be valid: {error}"));
        assert_eq!(sheet.index(), 3);
        assert_eq!(
            sheet.tables().map(Table::name).collect::<Vec<_>>(),
            ["First", "Second"]
        );

        let duplicate = TableBuilder::new("First", Dimensions::new(1, 1))
            .finish()
            .unwrap_or_else(|error| panic!("duplicate table should be valid: {error}"));
        assert!(matches!(
            Sheet::try_from_tables(
                "Sheet",
                0,
                vec![
                    TableBuilder::new("First", Dimensions::new(1, 1))
                        .finish()
                        .unwrap_or_else(|error| panic!("first table should be valid: {error}")),
                    duplicate,
                ],
            ),
            Err(Error::DuplicateTableName { name }) if name == "First"
        ));
    }
}
