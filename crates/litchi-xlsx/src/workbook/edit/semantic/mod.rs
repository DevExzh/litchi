//! Contextual semantic editing facade.
//!
//! The public edit API remains flat at the parent facade while the implementation
//! is organized by transaction, sheet handles, and grid layout semantics.

mod layout;
mod transaction;
mod worksheet;

pub use layout::{ColumnEdit, DefaultsEdit, RowEdit};
pub use transaction::Edit;
pub use worksheet::{NewSheet, TabEdit, WorksheetEdit};
