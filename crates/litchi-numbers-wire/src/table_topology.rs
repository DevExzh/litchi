//! Payload rewrites of one table-axis insertion or deletion, shared by the
//! Numbers, Pages and Keynote table owners (record 0370).
//!
//! A TST table's axis edit rewrites, payload by payload:
//!
//! - the stored rows of every storage tile ([`storage::tile_rows`](crate::table_topology::storage::tile_rows)) or each
//!   row's current and pre-BNC cell-offset tables
//!   ([`storage::tile_columns`](crate::table_topology::storage::tile_columns)), so cell payloads keep their bytes;
//! - the row or column header buckets ([`storage::header_bucket`](crate::table_topology::storage::header_bucket));
//! - the row/column UID map ([`storage::UidMap`](crate::table_topology::storage::UidMap));
//! - the stroke sidecar's border layers ([`storage::strokes`](crate::table_topology::storage::strokes));
//! - the `TST.TableDataList` entries the removed cells referenced
//!   ([`lists::release`](crate::table_topology::lists::release));
//! - a table formula owner's table extents ([`extents::rewrite_extents`](crate::table_topology::extents::rewrite_extents)).
//!
//! Each function maps one native payload to its rewritten form (or `None`
//! when it is unchanged); fields a rewrite does not own keep their exact
//! bytes. The caller owns object lookup, identifier allocation, component
//! placement and publication, as with [`crate::table_create`]. Moved from
//! `litchi-numbers`' topology transaction (records 0309, 0318, 0329), which
//! ported them from the host editor's `row_insert`, `column_insert`,
//! `table_delete` and `stroke_layers` phases.

pub mod extents;
pub mod lists;
pub mod storage;
mod wire;

/// A typed failure of a topology payload rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The payload is malformed or ambiguous.
    Invalid,
    /// The payload is well formed but uses a shape these rewrites do not own
    /// (a row that would move between tiles, a hidden or bordered removed
    /// axis, a UID map they would not re-encode byte for byte, ...).
    Unsupported,
    /// A bounded allocation failed.
    Allocation {
        /// The requested amount.
        amount: usize,
    },
}

impl Error {
    fn from_wire(error: &litchi_iwa_common::Error) -> Self {
        match error {
            litchi_iwa_common::Error::Allocation { amount, .. } => {
                Self::Allocation { amount: *amount }
            },
            _ => Self::Invalid,
        }
    }
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Invalid => formatter.write_str("table topology payload is invalid"),
            Self::Unsupported => formatter.write_str("table topology payload is not supported"),
            Self::Allocation { amount } => {
                write!(
                    formatter,
                    "could not allocate {amount} table topology items"
                )
            },
        }
    }
}

impl std::error::Error for Error {}

/// Result of a topology payload rewrite.
pub type Result<T> = std::result::Result<T, Error>;

/// Reserve `amount` more slots in `vector`, or fail with
/// [`Error::Allocation`].
fn reserve<T>(vector: &mut Vec<T>, amount: usize) -> Result<()> {
    vector
        .try_reserve(amount)
        .map_err(|_error| Error::Allocation { amount })
}

/// The identifier of one `TSP.Reference` payload.
fn reference_identifier(payload: &[u8]) -> Result<u64> {
    use crate::table_write::WriteError;
    crate::table_create::reference_identifier(payload).map_err(|error| match error {
        WriteError::Unsupported => Error::Unsupported,
        WriteError::Allocation { requested } => Error::Allocation { amount: requested },
        _ => Error::Invalid,
    })
}
