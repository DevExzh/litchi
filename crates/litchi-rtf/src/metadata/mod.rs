//! Passive document metadata and external identities.

pub(crate) mod custom_xml;
pub(crate) mod data_store;
pub(crate) mod document_origin;
pub(crate) mod document_variable;
pub(crate) mod external_reference;
pub(crate) mod file_table;
pub(crate) mod generator;
pub(crate) mod info;
pub(crate) mod mail_merge;
pub(crate) mod theme;
pub(crate) mod user_property;
pub(crate) mod window_caption;
pub(crate) mod write_reservation;
pub(crate) mod xml_namespace;
pub(crate) mod xsl_transform;

pub use info::{
    DocumentInfo as Info, DocumentProtection as Protection, ProtectionLevel, ProtectionType,
    RtfTimestamp as Timestamp,
};
pub use user_property::{
    UserProperty, UserPropertyDateTime as DateTime, UserPropertyType as PropertyType,
    UserPropertyValue as Value,
};
