#![expect(
    clippy::module_name_repetitions,
    reason = "public names retain established OOXML facade terminology"
)]
//! Layered, inert `WordprocessingML` mail-merge owner.

mod adapter;
mod codec;
mod edit;
mod limit;
mod model;
mod package;
#[allow(
    dead_code,
    unreachable_pub,
    reason = "the detached typed draft is exercised by unit tests; the family stages through the package"
)]
mod transaction;

#[cfg(test)]
mod tests;

pub(crate) use adapter::{
    extract_recipients, is_mail_merge_relationship_type, is_settings_relationship, map_docx_error,
    validate_mail_merge_relationships, validate_mail_merge_relationships_in,
};
pub use codec::parse_settings_mail_merge;
pub(crate) use codec::{
    parse_settings_mail_merge_processed, parse_settings_mail_merge_with_mce_limits,
};
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
pub(crate) use limit::mce_limits_for_source;
pub use model::{
    Conformance, DataSourceObject, DataType, Destination, FieldMap, FieldMappingType,
    MainDocumentType, Recipient, Recipients, RelationshipId, Settings,
};
pub use package::{Source, Target};
pub use transaction::{Revision, Snapshot};

/// Content type for the inert mail-merge recipient-data part.
pub const RECIPIENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.mailMergeRecipientData+xml";
