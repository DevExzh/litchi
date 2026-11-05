//! Exact-source transactions for moving one rooted table between sheets.
//!
//! The relocation boundary is selector-first: callers identify the source
//! sheet, one table within that sheet, and an existing destination sheet. The
//! transaction owns all native graph and archive rewriting; object
//! identifiers, protobuf payloads, and package members are not part of the
//! public API.

/// Exact-source table-relocation transaction values.
pub use crate::package::table_relocation::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
};
