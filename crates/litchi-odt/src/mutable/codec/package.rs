//! Package/state transition seam for mutable XML snapshots.

use super::super::model::MutableDocument;
use super::xml;
use litchi_core::Result;

pub(super) fn with_content_xml<T>(
    document: &MutableDocument,
    operation: impl FnOnce(&str) -> Result<T>,
) -> Result<T> {
    xml::with_content_snapshot(document, operation)
}

pub(super) fn invalidate_content_xml(document: &mut MutableDocument) {
    document.content_xml = None;
}
