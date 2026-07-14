//! Numbers object-reference extraction.

use super::super::ObjectIndex;
use crate::archive::RawMessage;
use prost::Message;

impl ObjectIndex {
    pub(super) fn extract_numbers_references(&mut self, object_id: u64, message: &RawMessage) {
        if let Ok(sheet) = crate::protobuf::tn::SheetArchive::decode(&*message.data) {
            for drawable in &sheet.drawable_infos {
                self.extract_reference(object_id, drawable);
            }
            if let Some(ref header) = sheet.header_storage {
                self.extract_reference(object_id, header);
            }
            if let Some(ref footer) = sheet.footer_storage {
                self.extract_reference(object_id, footer);
            }
        }
    }
}
