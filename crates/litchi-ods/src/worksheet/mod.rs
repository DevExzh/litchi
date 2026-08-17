//! Semantic ODS worksheet facade.
//!
//! The public surface is intentionally small and contextual: callers work
//! with [`Sheet`], [`Row`], [`Cell`], and [`CellValue`] while the package,
//! transaction, validation, and XML codec layers remain behind this facade.
//! Repeated rows and cells stay as physical runs, so logical lookup never
//! requires expanding a large ODF repetition into heap objects.

pub mod model;

pub(crate) mod codec;
pub(crate) mod package;
pub(crate) mod snapshot;
pub(crate) mod transaction;
pub(crate) mod validation;

pub use model::{Cell, CellValue, CellView, Merge, Row, Sheet};
pub use snapshot::{CellChange, Commit, Edit, MAX_CELL_CHANGES, Patch, Selector, Snapshot};
