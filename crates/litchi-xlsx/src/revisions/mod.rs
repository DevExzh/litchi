//! Layered `SpreadsheetML` workbook revision ownership.
//!
//! Revision values and invariants live in the model module, bounded XML in
//! the codec module, and workbook/revisionLog OPC graph operations in package.
//!
//! [`crate::Package::edit_revisions`] starts an [`Edit`] of the workbook's
//! revision owner; [`Edit::commit`] returns a [`Commit`] holding the new
//! package, the exact [`Patch`] and the [`Diagnostics`], and
//! [`crate::Package::apply_revisions`] replays a patch. The source package
//! never changes.

mod codec;
mod edit;
mod model;
mod package;
mod patch;
mod snapshot;
pub mod validation;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

pub use codec::{
    parse_revision_headers, parse_revision_log, parse_revision_users, write_revision_headers,
    write_revision_log, write_revision_users,
};
#[cfg(test)]
pub(crate) use edit::Transaction;
pub use edit::{Edit, EditOptions};
pub use model::{
    RevisionAttribute, RevisionAttributeNamespace, RevisionConformance, RevisionHeader,
    RevisionHeaderProperties, RevisionHeaders, RevisionLog, RevisionLogPart, RevisionRecord,
    RevisionRecordKind, RevisionUser, RevisionUsers, RevisionXmlElement, Revisions,
};
pub use package::load_workbook_revisions;
#[cfg(test)]
pub(crate) use package::store_workbook_revisions;
pub use patch::{Commit, Diagnostics, Error, LimitKind, Patch};
pub use snapshot::Snapshot;

#[cfg(test)]
mod tests;
