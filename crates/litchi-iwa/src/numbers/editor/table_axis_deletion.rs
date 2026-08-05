//! IWA adapter names for section-relative table topology edits.
//!
//! The semantic values are owned by `litchi-numbers`; this module keeps the
//! archive editor's existing focused imports while native resolution and wire
//! mutation remain in this crate.

pub use litchi_numbers::table::topology::{
    ColumnDeletion as TableColumnDeletion, RowDeletion as TableRowDeletion,
};
