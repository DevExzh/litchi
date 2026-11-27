//! Layered BIFF8 list-object semantic model.
//!
//! The facade keeps the historical public surface stable while the sibling
//! modules separate identity, style, source, opaque, table, and validation
//! concerns.

mod identity;
mod opaque;
mod source;
mod style;
mod table;
mod validation;

pub use identity::{ListColumnId, ListObjectId, ListObjectRange, ListTotalAggregation};
pub use opaque::{OpaqueListObjectFeature, OpaqueListObjectFutureRecord};
pub use source::{
    CachedDiskHeader, ExternalTableField, ExternalTableMetadata, ExternalTableVersion,
    ListObjectFeatureVersion, ListObjectSourceMetadata, WebColumnType, WebDefaultValue,
    WebEditMode, WebFieldInfo, WebInvalidCell, WebReadingOrder, WebTableField, WebTableMetadata,
    XmlColumnMapping, XmlDataType, XmlTableField, XmlTableMetadata,
};
pub use style::{ListObjectColumn, ListObjectStyleOptions};
pub use table::ListObject;

pub(in crate::list_object) use validation::{
    validate_column_name, validate_name, validate_table_name,
};
