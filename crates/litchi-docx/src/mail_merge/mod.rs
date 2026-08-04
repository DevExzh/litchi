//! Layered, inert WordprocessingML mail-merge owner.

mod codec;
mod model;
mod package;

pub use codec::parse_settings_mail_merge;
pub use model::{
    Conformance, DataSourceObject, DataType, Destination, FieldMap, FieldMappingType,
    MainDocumentType, Recipient, Recipients, Settings,
};
pub use package::{Source, Target};

/// Content type for the inert mail-merge recipient-data part.
pub const RECIPIENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.mailMergeRecipientData+xml";

// Historical names remain aliases so every existing API path keeps the same type.
pub type MailMergeSource = Source;
pub type MailMergeTarget = Target;
pub type MailMergeConformance = Conformance;
pub type MailMergeMainDocumentType = MainDocumentType;
pub type MailMergeDataType = DataType;
pub type MailMergeDestination = Destination;
pub type MailMergeFieldMappingType = FieldMappingType;
pub type MailMergeFieldMap = FieldMap;
pub type MailMergeDataSourceObject = DataSourceObject;
pub type MailMergeSettings = Settings;
pub type MailMergeRecipient = Recipient;
pub type MailMergeRecipients = Recipients;
