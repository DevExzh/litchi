//! Shared legacy Office Custom XML data-store semantics.
//!
//! The item payload is structurally validated and retained verbatim. Schema
//! references are metadata only: this module never resolves a URI, loads a
//! schema, expands an external entity, or executes application-specific XML.

mod codec;
mod model;
mod patch;
mod snapshot;
mod transaction;
mod xml;

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "tests use concise assertions while exercising fallible malformed-input paths"
)]
mod tests;

pub use self::codec::{inspect, inspect_with_limits, parse_properties, write, write_properties};
pub use self::model::{
    Error, Item, ItemId, ItemKind, Limits, Promotion, Properties, Result, RootName, Store,
};
pub use self::patch::{Change, Patch};
pub use self::snapshot::{Revision, Snapshot};
pub use self::transaction::{Commit, Transaction, update};

pub const STORE_STORAGE: &str = "MsoDataStore";
pub const REDUNDANT_PROMOTION_STORAGE: &str = "IsRedundantDataStorePromotion";
pub const MODIFIED_PROMOTION_STORAGE: &str = "IsModifiedDataStorePromotion";
pub const ITEM_STREAM: &str = "Item";
pub const PROPERTIES_STREAM: &str = "Properties";
pub const CUSTOM_XML_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/customXml";

pub(crate) const CUSTOM_PROPERTY_EDITOR_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/2006/customDocumentInformationPanel";
pub(crate) const CUSTOM_XSN_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/2006/metadata/customXsn";
pub(crate) const CONTENT_TYPE_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/2006/metadata/contentType";
pub(crate) const COVER_PAGE_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/2006/coverPageProps";
pub(crate) const LONG_PROPERTIES_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/2006/metadata/longProperties";
pub(crate) const CAML_NAMESPACE: &str = "office.server.policy";
