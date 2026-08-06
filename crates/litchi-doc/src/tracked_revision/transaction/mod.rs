//! Source-checked snapshot transactions for typed DOC tracked revisions.
//!
//! The transaction layer owns the complete DOC source artifact while the
//! package layer continues to own the Word FIB, piece, FKP, PLC, and SPRM
//! rewrites. Keeping these layers separate makes the public facade semantic:
//! callers edit revision metadata and ranges, not CFB streams or FKP pages.

mod commit;
mod edit;
mod patch;
mod snapshot;

pub use commit::Commit;
pub use edit::Transaction;
pub use patch::Patch;
pub use snapshot::Snapshot;

use crate::package::Error as PackageError;

/// Errors produced by a staged tracked-revision edit or source-checked patch.
#[derive(Debug)]
pub enum TransactionError {
    /// The candidate violates a DOC, CFB, resource, or SPRM invariant.
    Invalid(PackageError),
    /// A patch was applied to a snapshot other than its exact source.
    Conflict,
}

impl std::fmt::Display for TransactionError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::Conflict => formatter.write_str("tracked-revision transaction source conflict"),
        }
    }
}

impl std::error::Error for TransactionError {}

impl From<PackageError> for TransactionError {
    fn from(error: PackageError) -> Self {
        Self::Invalid(error)
    }
}

type Result<T> = std::result::Result<T, TransactionError>;
