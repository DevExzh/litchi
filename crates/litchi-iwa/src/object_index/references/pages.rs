//! Application-specific object-reference extraction.

use super::super::ObjectIndex;
use crate::archive::RawMessage;

impl ObjectIndex {
    pub(super) fn extract_pages_references(&mut self, object_id: u64, raw_msg: &RawMessage) {
        use prost::Message;

        match raw_msg.type_ {
            // TP (Pages) types
            10000 => {
                // TP.DocumentArchive
                if let Ok(doc) = crate::protobuf::tp::DocumentArchive::decode(&*raw_msg.data) {
                    // Extract theme reference
                    if let Some(ref theme) = doc.theme {
                        self.extract_reference(object_id, theme);
                    }

                    // Extract stylesheet reference
                    if let Some(ref stylesheet) = doc.stylesheet {
                        self.extract_reference(object_id, stylesheet);
                    }
                }
            },

            10011 => {
                // TP.SectionArchive
                // Note: SectionArchive has a complex structure
                // References are embedded in nested structures
            },

            _ => {},
        }
    }
}
