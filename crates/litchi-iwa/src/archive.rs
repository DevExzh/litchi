//! Application-facing IWA archive facade.
//!
//! Neutral archive parsing, metadata preservation, bounded mutation, and
//! serialization are owned by [`litchi_iwa_archive::iwa`]. This module keeps
//! the format crate's established import path and adds only the
//! application-level decoded-message text projection.

use crate::protobuf::decode_common;

pub use litchi_iwa_archive::iwa::{
    Archive, ArchiveLimits, ArchiveObject, ArchiveReferenceOccurrence, ArchiveReferencePolicy,
    ArchiveReferenceVisitor, FieldInfo, FieldObjectReferenceTransition, FieldPath, FieldType,
    MessageInfo, ObjectReferenceTransition, RawMessage, UnknownFieldRule,
};

// Keep the result name local to the migration host while retaining the exact
// neutral-core type. This alias is intentionally crate-visible: callers should
// use the typed format APIs, while the remaining host adapters need to return
// the shared physical-layer result without importing the retired direct core
// dependency.
pub(crate) use litchi_iwa_archive::iwa::Result as CoreResult;

// These names are used only by host unit tests when matching bounded physical
// errors. Keep them out of production builds so the migration host remains
// warning-free under the strict `-D warnings` lint gate.
#[cfg(test)]
pub(crate) use litchi_iwa_archive::iwa::{Error as CoreError, LimitKind as CoreLimitKind};

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
