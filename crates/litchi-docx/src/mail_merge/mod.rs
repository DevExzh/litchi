//! Layered, inert WordprocessingML mail-merge owner.

mod adapter;
mod codec;
mod model;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub(crate) use adapter::{
    extract_recipients, is_mail_merge_relationship_type, is_settings_relationship, map_docx_error,
    validate_mail_merge_relationships,
};
pub use codec::parse_settings_mail_merge;
pub use model::{
    Conformance, DataSourceObject, DataType, Destination, FieldMap, FieldMappingType,
    MainDocumentType, Recipient, Recipients, Settings,
};
pub use package::{Source, Target};
pub use transaction::{Commit, Patch, Revision, Snapshot, Transaction};

/// Content type for the inert mail-merge recipient-data part.
pub const RECIPIENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.mailMergeRecipientData+xml";
