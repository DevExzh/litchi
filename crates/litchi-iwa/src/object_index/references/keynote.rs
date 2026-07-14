//! Application-specific object-reference extraction.

use super::super::ObjectIndex;
use crate::archive::RawMessage;

impl ObjectIndex {
    pub(super) fn extract_keynote_references(&mut self, object_id: u64, raw_msg: &RawMessage) {
        use prost::Message;

        match raw_msg.type_ {
            // KN (Keynote) types
            5 | 6 => {
                // KN.SlideArchive contains references to drawables, builds, and transitions
                if let Ok(slide) = crate::protobuf::kn::SlideArchive::decode(&*raw_msg.data) {
                    // Extract style reference
                    self.extract_reference(object_id, &slide.style);

                    // Extract drawable references (shapes, images, text boxes)
                    for drawable in &slide.drawables {
                        self.extract_reference(object_id, drawable);
                    }

                    // Extract build animation references
                    for build in &slide.builds {
                        self.extract_reference(object_id, build);
                    }

                    // Extract placeholder references
                    if let Some(ref title) = slide.title_placeholder {
                        self.extract_reference(object_id, title);
                    }
                    if let Some(ref body) = slide.body_placeholder {
                        self.extract_reference(object_id, body);
                    }
                    if let Some(ref object) = slide.object_placeholder {
                        self.extract_reference(object_id, object);
                    }
                    if let Some(ref slide_num) = slide.slide_number_placeholder {
                        self.extract_reference(object_id, slide_num);
                    }

                    // Extract style references
                    for para_style in &slide.body_paragraph_styles {
                        self.extract_reference(object_id, para_style);
                    }
                    for list_style in &slide.body_list_styles {
                        self.extract_reference(object_id, list_style);
                    }
                }
            },

            2 => {
                // KN.ShowArchive (conflicts with TSP.MessageInfo, handle by context)
                // Try to decode as ShowArchive for Keynote documents
                if let Ok(show) = crate::protobuf::kn::ShowArchive::decode(&*raw_msg.data) {
                    // Extract theme and stylesheet references
                    self.extract_reference(object_id, &show.theme);
                    self.extract_reference(object_id, &show.stylesheet);

                    // Extract UI state reference
                    if let Some(ref ui_state) = show.ui_state {
                        self.extract_reference(object_id, ui_state);
                    }

                    // Extract recording reference if present
                    if let Some(ref recording) = show.recording {
                        self.extract_reference(object_id, recording);
                    }

                    // Note: Slide references are in the slide_tree structure
                    // which is not a simple Reference type
                }
            },

            _ => {},
        }
    }
}
