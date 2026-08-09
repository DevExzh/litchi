//! Layered `SpreadsheetML` workbook revision ownership.
//!
//! Revision values and invariants live in the model module, bounded XML in
//! the codec module, and workbook/revisionLog OPC graph operations in package.

mod codec;
mod model;
mod package;
pub mod patch;
pub mod snapshot;
pub mod transaction;
pub mod validation;

use crate::error::Error;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn limit(name: &str) -> Error {
    invalid(format!("{name} exceeds configured limit"))
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

pub use codec::{
    parse_revision_headers, parse_revision_log, parse_revision_users, write_revision_headers,
    write_revision_log, write_revision_users,
};
pub use model::{
    RevisionAttribute, RevisionAttributeNamespace, RevisionConformance, RevisionHeader,
    RevisionHeaderProperties, RevisionHeaders, RevisionLog, RevisionLogPart, RevisionRecord,
    RevisionRecordKind, RevisionUser, RevisionUsers, RevisionXmlElement, Revisions,
};
pub use package::{load_workbook_revisions, remove_workbook_revisions, store_workbook_revisions};
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
pub use transaction::Transaction;

#[cfg(test)]
mod tests;
