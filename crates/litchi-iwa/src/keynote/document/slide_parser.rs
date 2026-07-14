//! Slide protobuf decoding and referenced-content extraction.

use super::KeynoteDocument;
use crate::Result;
use crate::keynote::slide::{self, KeynoteSlide};
use crate::text::TextExtractor;

impl KeynoteDocument {
    /// Parse a single slide from an object
    pub(super) fn parse_slide(
        &self,
        index: usize,
        object: &crate::archive::ArchiveObject,
    ) -> Result<KeynoteSlide> {
        use prost::Message;

        let mut slide = KeynoteSlide::new(index);

        // Extract text content from the slide object
        let text_parts = object.extract_text();

        if !text_parts.is_empty() {
            // First text part is typically the title or slide name
            slide.title = text_parts.first().cloned();

            // Remaining parts are content
            slide.text_content = text_parts.into_iter().skip(1).collect();
        }

        // Parse the SlideArchive protobuf message
        // KN.SlideArchive contains:
        // - name: string (slide title)
        // - note: reference to KN.NoteArchive (speaker notes)
        // - drawables: references to drawable objects (shapes, text boxes, images)
        // - builds: references to KN.BuildArchive (animations)
        // - transition: TransitionArchive (transition effect)
        // - master: reference to master slide

        if let Some(raw_message) = object.messages.first() {
            // Try to decode as SlideArchive
            if let Ok(slide_archive) = crate::protobuf::kn::SlideArchive::decode(&*raw_message.data)
            {
                // Extract slide name if available
                if let Some(ref name) = slide_archive.name
                    && !name.is_empty()
                {
                    slide.title = Some(name.clone());
                }

                // Extract master slide reference
                if let Some(ref master) = slide_archive.master {
                    slide.master_slide_id = Some(master.identifier);
                }

                // Extract build animations
                for build_ref in &slide_archive.builds {
                    if let Ok(build) = self.extract_build_animation(build_ref.identifier) {
                        slide.builds.push(build);
                    }
                }

                // Extract transition
                slide.transition = self.parse_transition(&slide_archive.transition);

                // Resolve drawable references to get text boxes and other content
                for drawable_ref in &slide_archive.drawables {
                    if let Ok(text_content) = self.extract_drawable_text(drawable_ref.identifier)
                        && !text_content.is_empty()
                    {
                        slide.text_content.push(text_content);
                    }
                }

                // Extract speaker notes
                if let Some(ref note_ref) = slide_archive.note
                    && let Ok(notes) = self.extract_speaker_notes(note_ref.identifier)
                {
                    slide.notes = Some(notes);
                }
            }
        }

        // Extract text from text storages
        let extractor = TextExtractor::new();
        if let Ok(storage) = extractor.extract_from_object(object)
            && !storage.is_empty()
        {
            slide.text_storages.push(storage);
        }

        Ok(slide)
    }

    /// Extract build animation from a BuildArchive object
    fn extract_build_animation(&self, build_id: u64) -> Result<slide::BuildAnimation> {
        use prost::Message;
        use slide::{BuildAnimation, BuildAnimationType};

        if let Some(resolved) = self.object_index.resolve_object(&self.bundle, build_id)? {
            for msg in &resolved.messages {
                if let Ok(build_archive) = crate::protobuf::kn::BuildArchive::decode(&*msg.data) {
                    let animation_type = Self::parse_build_delivery(&build_archive.delivery);
                    let target_id = Some(build_archive.drawable.identifier);
                    let duration = build_archive.duration as f32;

                    return Ok(BuildAnimation {
                        animation_type,
                        target_id,
                        duration,
                    });
                }
            }
        }

        // Return a default build if parsing failed
        Ok(BuildAnimation {
            animation_type: BuildAnimationType::Other,
            target_id: None,
            duration: 0.0,
        })
    }

    /// Parse build delivery string into animation type
    fn parse_build_delivery(delivery: &str) -> slide::BuildAnimationType {
        use slide::BuildAnimationType;

        match delivery.to_lowercase().as_str() {
            s if s.contains("appear") => BuildAnimationType::Appear,
            s if s.contains("dissolve") => BuildAnimationType::Dissolve,
            s if s.contains("move") => BuildAnimationType::MoveIn,
            s if s.contains("scale") && s.contains("fade") => BuildAnimationType::FadeAndScale,
            s if s.contains("scale") => BuildAnimationType::Scale,
            _ => BuildAnimationType::Other,
        }
    }

    /// Parse transition archive into slide transition
    fn parse_transition(
        &self,
        transition: &crate::protobuf::kn::TransitionArchive,
    ) -> Option<slide::SlideTransition> {
        use slide::{SlideTransition, TransitionType};

        // Extract duration from attributes
        // The attributes field is required (not Optional)
        let duration = transition.attributes.database_duration.unwrap_or(0.0) as f32;

        // Determine transition type from attributes
        // The actual transition type is embedded in the attributes structure
        // For now, we use a generic transition type
        let transition_type = TransitionType::Other;

        Some(SlideTransition {
            transition_type,
            duration,
        })
    }

    /// Extract text content from a drawable object
    fn extract_drawable_text(&self, drawable_id: u64) -> Result<String> {
        use prost::Message;

        if let Some(resolved) = self
            .object_index
            .resolve_object(&self.bundle, drawable_id)?
        {
            // Drawables can contain text storages
            for msg in &resolved.messages {
                // Try to extract text from TSWP storage messages (types 2001-2022)
                if msg.type_ >= 2001
                    && msg.type_ <= 2022
                    && let Ok(storage) = crate::protobuf::tswp::StorageArchive::decode(&*msg.data)
                    && !storage.text.is_empty()
                {
                    return Ok(storage.text.join(" "));
                }
            }

            // Also try generic text extraction from the resolved object
            for msg in &resolved.messages {
                if let Ok(storage) = crate::protobuf::tswp::StorageArchive::decode(&*msg.data)
                    && !storage.text.is_empty()
                {
                    return Ok(storage.text.join(" "));
                }
            }
        }

        Ok(String::new())
    }

    /// Extract speaker notes from a NoteArchive object
    fn extract_speaker_notes(&self, note_id: u64) -> Result<String> {
        use prost::Message;

        if let Some(resolved) = self.object_index.resolve_object(&self.bundle, note_id)? {
            for msg in &resolved.messages {
                if let Ok(note_archive) = crate::protobuf::kn::NoteArchive::decode(&*msg.data) {
                    // The note contains a reference to a TSWP.StorageArchive
                    let storage_id = note_archive.contained_storage.identifier;
                    if let Some(storage_obj) =
                        self.object_index.resolve_object(&self.bundle, storage_id)?
                    {
                        for storage_msg in &storage_obj.messages {
                            if let Ok(storage) =
                                crate::protobuf::tswp::StorageArchive::decode(&*storage_msg.data)
                            {
                                let notes_text = storage.text.join("\n");
                                if !notes_text.is_empty() {
                                    return Ok(notes_text);
                                }
                            }
                        }
                    }
                }
            }
        }

        Ok(String::new())
    }
}
