//! Layered BIFF8 worksheet-table owner.
//!
//! Semantic table values and validation live in `model`, BIFF payload codecs
//! live in `codec`, worksheet record-sequence assembly lives in `package`, and
//! focused regression coverage lives in `tests`. Unsupported table and future
//! records remain bounded and opaque so they can be written back losslessly.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub(crate) const FEAT_HDR11_RECORD_TYPE: u16 = 0x0871;
pub(crate) const FEATURE11_RECORD_TYPE: u16 = 0x0872;
pub(crate) const CONTINUE_FRT11_RECORD_TYPE: u16 = 0x0875;
pub(crate) const LIST12_RECORD_TYPE: u16 = 0x0877;
pub(crate) const FEATURE12_RECORD_TYPE: u16 = 0x0878;
pub(crate) const AUTO_FILTER12_RECORD_TYPE: u16 = 0x087e;

pub(super) const ISF_LIST: u16 = 5;
pub(super) const MAX_PAYLOAD: usize = 8_224;
pub(super) const MAX_CONTINUE_RGB: usize = 8_212;
pub(super) const MAX_FEATURE_BYTES: usize = 1_048_576;

pub(super) fn invalid(rt: u16, message: impl Into<String>) -> crate::Error {
    crate::Error::InvalidRecord {
        record_type: rt,
        message: message.into(),
    }
}

pub use model::{
    CachedDiskHeader, ExternalTableField, ExternalTableMetadata, ExternalTableVersion,
    ListColumnId, ListObject, ListObjectColumn, ListObjectFeatureVersion, ListObjectId,
    ListObjectRange, ListObjectSourceMetadata, ListObjectStyleOptions, ListTotalAggregation,
    OpaqueListObjectFeature, OpaqueListObjectFutureRecord, WebColumnType, WebDefaultValue,
    WebEditMode, WebFieldInfo, WebInvalidCell, WebReadingOrder, WebTableField, WebTableMetadata,
    XmlColumnMapping, XmlDataType, XmlTableField, XmlTableMetadata,
};

pub(crate) use package::{ListObjectCollector, feature_header_record};
