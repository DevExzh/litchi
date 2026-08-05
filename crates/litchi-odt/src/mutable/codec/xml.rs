//! XML snapshot primitives used by mutable semantic operations.

use super::super::model::MutableDocument;
use litchi_core::Result;

/// Apply a read-only operation to the current content snapshot.
pub(super) fn with_content_snapshot<T>(
    document: &MutableDocument,
    operation: impl FnOnce(&str) -> Result<T>,
) -> Result<T> {
    if let Some(xml) = document.content_xml.as_deref() {
        operation(xml)
    } else {
        let xml = document.generate_content_xml();
        operation(&xml)
    }
}
