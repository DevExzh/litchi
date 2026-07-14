//! Shared text object-reference extraction.

use super::super::ObjectIndex;
use crate::archive::RawMessage;
use prost::Message;

impl ObjectIndex {
    pub(super) fn extract_text_references(&mut self, object_id: u64, message: &RawMessage) {
        if let Ok(storage) = crate::protobuf::tswp::StorageArchive::decode(&*message.data)
            && let Some(ref style_sheet) = storage.style_sheet
        {
            self.extract_reference(object_id, style_sheet);
        }
    }
}
