//! Section-relative table row and column insertion and removal (record
//! 0309): the atomic, exact-source [`Edit`](crate::table::topology::Edit) started by
//! [`crate::Package::edit_table_topology`], and the section-relative
//! coordinates it takes.
//!
//! The coordinate values describe semantic table positions only; the
//! transaction resolves them against the table's native header, body and
//! footer counts and rewrites storage, headers, UIDs, borders, counts and
//! every formula coordinate that addresses the table (record 0318: the
//! table's own formulas and other tables' formulas that read it) in one
//! publication.
//!
//! The coordinate values are the neutral `litchi_iwa_common::table::topology`
//! vocabulary (record 0368), re-exported here.

pub use crate::package::table_topology::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal, Unsupported,
};
pub use litchi_iwa_common::table::topology::{
    ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion,
};
