//! Layered, inert WordprocessingML mail-merge owner.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::parse_settings_mail_merge;
pub use model::{
    Conformance, DataSourceObject, DataType, Destination, FieldMap, FieldMappingType,
    MainDocumentType, Recipient, Recipients, Settings,
};
pub use package::{Source, Target};

/// Content type for the inert mail-merge recipient-data part.
pub const RECIPIENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.mailMergeRecipientData+xml";
