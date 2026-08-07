//! Archive-bound Numbers sheet conversion.
//!
//! The semantic sheet owner lives in `litchi_numbers::sheet`. This module is
//! deliberately crate-private: it retains only the temporary mutable state
//! needed while decoding native archives and transfers finished tables to the
//! immutable leaf model at the boundary.

use super::table::{Table, map_table_error};

/// Archive adapter used while decoding one native sheet.
///
/// This type is intentionally opaque outside the Numbers package adapter;
/// callers that need archive-free semantics should use [`crate::Document`].
#[derive(Debug, Clone)]
pub(super) struct DecodedSheet {
    pub(crate) name: String,
    pub(crate) index: usize,
    pub(crate) tables: Vec<Table>,
}

impl DecodedSheet {
    /// Creates a native sheet adapter.
    pub(super) fn new(name: String, index: usize) -> Self {
        Self {
            name,
            index,
            tables: Vec::new(),
        }
    }

    /// Adds an archive-decoded table without relying on infallible growth.
    pub(super) fn try_add_table(&mut self, table: Table) -> super::Result<()> {
        self.tables.try_reserve(1).map_err(|_error| {
            super::Error::Common(litchi_iwa_common::Error::Allocation {
                resource: "Numbers rooted sheet tables",
                amount: self.tables.len().saturating_add(1),
            })
        })?;
        self.tables.push(table);
        Ok(())
    }

    /// Borrows the native sheet name without exposing archive state.
    #[must_use]
    pub(super) fn name(&self) -> &str {
        &self.name
    }

    /// Returns the logical zero-based sheet position.
    #[must_use]
    pub(super) const fn index(&self) -> usize {
        self.index
    }

    /// Iterates native tables without exposing the backing collection.
    #[must_use]
    pub(super) fn tables(&self) -> impl ExactSizeIterator<Item = &Table> + '_ {
        self.tables.iter()
    }

    /// Returns a native table by exact name.
    #[must_use]
    pub(super) fn table(&self, name: &str) -> Option<&Table> {
        self.tables.iter().find(|table| table.name() == name)
    }

    /// Returns the number of native tables.
    pub(super) fn table_count(&self) -> usize {
        self.tables.len()
    }

    /// Returns the number of addressable cells across all native tables.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "Retained for native table-budget accounting once callers migrate to the format-owned adapter."
        )
    )]
    pub(super) fn total_cell_count(&self) -> usize {
        self.tables
            .iter()
            .map(|table| table.row_count() * table.column_count())
            .sum()
    }

    /// Consumes the archive adapter into the canonical dependency-free sheet.
    pub(super) fn into_semantic(self) -> super::Result<crate::Sheet> {
        let mut builder = crate::sheet::Builder::new(self.name, self.index);
        for table in self.tables {
            let semantic = table.into_semantic()?;
            builder
                .push_table(semantic)
                .map_err(|failure| map_table_error(failure.into_parts().0))?;
        }
        Ok(builder.finish())
    }
}
