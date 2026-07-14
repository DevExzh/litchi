//! Protobuf reference extraction while building an object index.

use super::ObjectIndex;
use crate::Result;
use crate::archive::ArchiveObject;

mod chart;
mod drawing;
mod keynote;
mod numbers;
mod pages;
mod table;
mod text;

impl ObjectIndex {
    pub(super) fn parse_object_references(
        &mut self,
        object_id: u64,
        object: &ArchiveObject,
    ) -> Result<()> {
        for message in &object.messages {
            match message.type_ {
                6000 | 6001 | 6005 | 6201 => self.extract_table_references(object_id, message),
                2001..=2022 => self.extract_text_references(object_id, message),
                2 | 5 | 6 => self.extract_keynote_references(object_id, message),
                3 => self.extract_numbers_references(object_id, message),
                3002..=3009 => self.extract_drawing_references(object_id, message),
                5000 | 5004 | 5020 | 5021 => self.extract_chart_references(object_id, message),
                10000 | 10011 => self.extract_pages_references(object_id, message),
                _ => {},
            }
        }
        Ok(())
    }

    fn extract_reference(&mut self, source_id: u64, reference: &crate::protobuf::tsp::Reference) {
        if reference.identifier != 0 {
            self.reference_graph
                .add_reference(source_id, reference.identifier);
        }
    }
}
