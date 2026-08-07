//! Application-facing IWA archive facade.
//!
//! Neutral archive parsing, metadata preservation, bounded mutation, and
//! serialization are owned by [`litchi_iwa_core`]. This module keeps the
//! format crate's established import path and adds only the application-level
//! decoded-message text projection.

use crate::protobuf::decode_common;

pub use litchi_iwa_core::{
    Archive, ArchiveInfo, ArchiveLimits, ArchiveObject, FieldInfo, FieldPath, FieldType,
    MessageInfo, RawMessage, UnknownFieldRule,
};

/// Extract application text from an archive object without storing decoded
/// protobuf trait objects in the neutral archive model.
pub(crate) fn extract_text(object: &ArchiveObject) -> Vec<String> {
    let mut text = Vec::new();
    for message in &object.messages {
        if let Ok(decoded) = decode_common(message.type_, &message.data) {
            text.extend(decoded.extract_text());
        }
    }
    text
}
