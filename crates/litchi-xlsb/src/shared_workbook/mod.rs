//! Layered, inert XLSB shared-workbook metadata ownership.
//!
//! `BrtInfo`, `BrtRRHeader`, and `BrtUsr` are exposed as typed metadata.  The
//! revision-log parts remain ordered opaque BIFF12 records: the owner never
//! applies, replays, locks, or collaborates on revisions.

pub mod codec;
pub mod model;
pub mod package;
pub mod transaction;
pub mod validation;

use crate::package::error::Error;

/// Result type used by shared-workbook package and transaction operations.
pub type Result<T> = std::result::Result<T, Error>;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(format!("XLSB shared-workbook: {}", message.into()))
}

pub(super) fn unsupported(message: impl Into<String>) -> Error {
    Error::UnsupportedFeature(format!("XLSB shared-workbook: {}", message.into()))
}

pub(super) fn map_raw(error: crate::raw::Error) -> Error {
    invalid(error.to_string())
}

pub use model::{
    Catalog, Guid, Header, Info, RawRecord, RecordView, RevisionEnvelope, RevisionHeaders,
    RevisionLog, ShortDateTime, User, UserNames,
};
pub use package::{load, store};
pub use transaction::{
    Commit, Patch, Snapshot, SourcePart, SourceRelationship, Transaction, apply, read,
};
pub use validation::{validate_catalog, validate_headers, validate_log, validate_users};

#[cfg(test)]
mod tests;
