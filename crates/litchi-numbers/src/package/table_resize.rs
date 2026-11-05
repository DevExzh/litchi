//! Implementation of the public [`crate::table::resize`] transaction:
//! resizing a table at its trailing edges (record 0370e).
//!
//! Growth adds blank trailing rows and columns; shrinking is accepted only
//! when the removed trailing region holds no stored cell. The tiles lose
//! the removed rows' records, the header buckets the removed rows' and
//! columns' headers, the row/column UID map its trailing identities (new
//! ones are allocated above every existing one), the stroke sidecar takes
//! the new size, and the table's own formula owner moves its table and body
//! extents. This is the host editor's `resize_table`; the staged rewrite is
//! the shared `litchi_iwa_table::topology::stage_resize` (record 0412g).

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use thiserror::Error as ThisError;

use super::Package;
use litchi_iwa_table::topology as engine;

use super::structure::{self, Published, Working, metadata::Transition};
use super::table_topology as topology;
use crate::selector::{SheetSelector, TableSelector};

structure::structural_family!("table-resize");

/// The most row and column identities one table may need.
const MAX_TABLE_UIDS: u64 = litchi_iwa_table::topology::MAX_TABLE_UIDS;

/// A typed, content-redacted resize failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted table source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} table items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested size.
    #[error("table resize candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("table resize patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The table selector did not resolve one table of the sheet.
    #[error("table selector did not resolve")]
    TableNotFound,
    /// Commit was requested before staging a size.
    #[error("no table size is staged")]
    NoStagedOperation,
    /// The size is zero, leaves no body row or column, or needs more row
    /// and column identities than a table may hold.
    #[error("the table cannot take this size")]
    InvalidSize,
    /// The removed trailing region holds a stored cell.
    #[error("the removed rows or columns hold stored cells")]
    StoredCells,
    /// The table's storage uses a layout this route does not rewrite.
    #[error("source representation does not support resizing this table")]
    UnsupportedSource,
    /// A surviving object depends on what the resize would change.
    #[error("a surviving object depends on the table's size")]
    UnsupportedDependency,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::TableNotFound => RefusalKind::NotFound,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::InvalidSize | Self::StoredCells => RefusalKind::InvalidValue,
            Self::UnsupportedSource | Self::UnsupportedDependency => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<topology::Error> for Error {
    fn from(error: topology::Error) -> Self {
        use topology::{Error as E, Refusal as R};
        match error {
            E::Refused(R::SheetNotFound) => Self::Refused(Refusal::SheetNotFound),
            E::Refused(R::TableNotFound) => Self::Refused(Refusal::TableNotFound),
            E::Refused(_) => Self::Refused(Refusal::UnsupportedSource),
            E::Allocation { amount } => Self::Allocation { amount },
            E::Verification => Self::Verification,
            _ => Self::InvalidSource,
        }
    }
}

impl From<engine::Error> for Error {
    fn from(error: engine::Error) -> Self {
        use engine::{Error as E, Refusal as R};
        match error {
            E::Refused(R::InvalidSize) => Self::Refused(Refusal::InvalidSize),
            E::Refused(R::StoredCells) => Self::Refused(Refusal::StoredCells),
            E::Refused(_) => Self::Refused(Refusal::UnsupportedSource),
            E::Fault(fault) => Self::from(fault),
            E::Allocation { amount } => Self::Allocation { amount },
            _ => Self::InvalidSource,
        }
    }
}

/// One immutable edit that resizes one table.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    table: usize,
    size: Option<(u32, u32)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("size", &self.size)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage `rows` × `columns` (header and footer rows and header columns
    /// included) as the table's size.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidSize`] for a zero size or one needing more row and
    /// column identities than a table may hold.
    pub fn resize(mut self, rows: u32, columns: u32) -> Result<Self, Error> {
        if rows == 0 || columns == 0 || u64::from(rows) + u64::from(columns) > MAX_TABLE_UIDS {
            return Err(Error::Refused(Refusal::InvalidSize));
        }
        self.size = Some((rows, columns));
        Ok(self)
    }

    /// Validate and atomically publish the staged size; the current size
    /// is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::StoredCells`],
    /// [`Refusal::InvalidSize`] for a size leaving no body row or column,
    /// source, resource and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        let (rows, columns) = self
            .size
            .ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        let Some(published) = resize(self.source, self.sheet, self.table, rows, columns)? else {
            return Commit::unchanged(self.source);
        };
        let table = published.package.workbook_ref().sheets()[self.sheet]
            .tables()
            .nth(self.table)
            .ok_or(Error::Verification)?;
        if (table.row_count(), table.column_count()) != (rows, columns) {
            return Err(Error::Verification);
        }
        Commit::new(self.source, published)
    }
}

fn resize(
    source: &Package,
    sheet: usize,
    table: usize,
    rows: u32,
    columns: u32,
) -> Result<Option<Published>, Error> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    let target = topology::locate(&working, sheet, table)?;
    if !engine::stage_resize(&mut working, &target, rows, columns)? {
        return Ok(None);
    }
    Ok(Some(working.publish(&Transition::default())?))
}

impl Package {
    /// Start an immutable resize of one table.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let commit = package.edit_table_resize(0, 0)?.resize(20, 6)?.commit()?;
    /// # let _ = commit;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector and source errors.
    pub fn edit_table_resize<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = self
            .workbook_ref()
            .sheet(sheet.into())
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?
            .index();
        let table = topology::table_position(self, sheet, table.into())?;
        Ok(Edit {
            source: self,
            sheet,
            table,
            size: None,
        })
    }

    /// Apply a table-resize patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_resize(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
