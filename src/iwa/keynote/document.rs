//! Keynote Document Implementation
//!
//! Provides high-level API for working with Apple Keynote presentations.

use std::path::Path;

use super::show::KeynoteShow;
use super::slide::KeynoteSlide;
use crate::iwa::Result;
use crate::iwa::bundle::Bundle;
use crate::iwa::object_index::ObjectIndex;
use crate::iwa::registry::Application;
use crate::iwa::text::TextExtractor;

/// High-level interface for Keynote documents
pub struct KeynoteDocument {
    /// Underlying bundle
    bundle: Bundle,
    /// Object index for cross-referencing
    object_index: ObjectIndex,
}

impl KeynoteDocument {
    /// Open a Keynote document from a path
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::keynote::KeynoteDocument;
    ///
    /// let doc = KeynoteDocument::open("presentation.key")?;
    /// println!("Loaded Keynote presentation");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let bundle = Bundle::open(path)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Ok(Self {
            bundle,
            object_index,
        })
    }

    /// Open a Keynote document from raw bytes
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::keynote::KeynoteDocument;
    /// use std::fs;
    ///
    /// let data = fs::read("presentation.key")?;
    /// let doc = KeynoteDocument::from_bytes(&data)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let bundle = Bundle::from_bytes(bytes)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Ok(Self {
            bundle,
            object_index,
        })
    }

    /// Create a Keynote document from raw bytes (ZIP archive data).
    ///
    /// This is used for single-pass parsing where the ZIP archive has already
    /// been validated during format detection. It avoids double-parsing.
    pub fn from_archive_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes)
    }

    /// Extract all text content from the presentation
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::keynote::KeynoteDocument;
    ///
    /// let doc = KeynoteDocument::open("presentation.key")?;
    /// let text = doc.text()?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        let mut extractor = TextExtractor::new();
        extractor.extract_from_bundle(&self.bundle)?;
        Ok(extractor.get_text())
    }

    /// Extract slides from the presentation
    ///
    /// Keynote presentations consist of slides with content, animations, and transitions.
    /// This method parses the presentation structure and returns all slides.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::keynote::KeynoteDocument;
    ///
    /// let doc = KeynoteDocument::open("presentation.key")?;
    /// let slides = doc.slides()?;
    ///
    /// for slide in slides {
    ///     println!("Slide {}", slide.index + 1);
    ///     if let Some(title) = &slide.title {
    ///         println!("  Title: {}", title);
    ///     }
    ///     for text in &slide.text_content {
    ///         println!("  - {}", text);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn slides(&self) -> Result<Vec<KeynoteSlide>> {
        let mut slides = Vec::new();

        // Find slide archives (message type 5/6 is KN.SlideArchive, type 1102 in our decoder)
        let slide_objects = self.bundle.find_objects_by_type(1102);

        if slide_objects.is_empty() {
            // Try alternate slide message types (5 and 6 from JSON)
            let alt_slide_objects_5 = self.bundle.find_objects_by_type(5);
            let alt_slide_objects_6 = self.bundle.find_objects_by_type(6);

            for (index, (_archive_name, object)) in alt_slide_objects_5
                .iter()
                .chain(alt_slide_objects_6.iter())
                .enumerate()
            {
                let slide = self.parse_slide(index, object)?;
                if !slide.is_empty() {
                    slides.push(slide);
                }
            }
        } else {
            for (index, (_archive_name, object)) in slide_objects.iter().enumerate() {
                let slide = self.parse_slide(index, object)?;
                if !slide.is_empty() {
                    slides.push(slide);
                }
            }
        }

        // If no slides found, create a default slide with all text
        if slides.is_empty() {
            let mut extractor = TextExtractor::new();
            extractor.extract_from_bundle(&self.bundle)?;

            if extractor.storage_count() > 0 {
                let mut slide = KeynoteSlide::new(0);
                for storage in extractor.storages() {
                    if !storage.is_empty() {
                        slide.text_storages.push(storage.clone());
                        slide.text_content.push(storage.plain_text().to_string());
                    }
                }
                if !slide.is_empty() {
                    slides.push(slide);
                }
            }
        }

        Ok(slides)
    }

    /// Parse a single slide from an object
    fn parse_slide(
        &self,
        index: usize,
        object: &crate::iwa::archive::ArchiveObject,
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
            if let Ok(slide_archive) =
                crate::iwa::protobuf::kn::SlideArchive::decode(&*raw_message.data)
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
    fn extract_build_animation(&self, build_id: u64) -> Result<super::slide::BuildAnimation> {
        use super::slide::{BuildAnimation, BuildAnimationType};
        use prost::Message;

        if let Some(resolved) = self.object_index.resolve_object(&self.bundle, build_id)? {
            for msg in &resolved.messages {
                if let Ok(build_archive) =
                    crate::iwa::protobuf::kn::BuildArchive::decode(&*msg.data)
                {
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
    fn parse_build_delivery(delivery: &str) -> super::slide::BuildAnimationType {
        use super::slide::BuildAnimationType;

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
        transition: &crate::iwa::protobuf::kn::TransitionArchive,
    ) -> Option<super::slide::SlideTransition> {
        use super::slide::{SlideTransition, TransitionType};

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
                    && let Ok(storage) =
                        crate::iwa::protobuf::tswp::StorageArchive::decode(&*msg.data)
                    && !storage.text.is_empty()
                {
                    return Ok(storage.text.join(" "));
                }
            }

            // Also try generic text extraction from the resolved object
            for msg in &resolved.messages {
                if let Ok(storage) = crate::iwa::protobuf::tswp::StorageArchive::decode(&*msg.data)
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
                if let Ok(note_archive) = crate::iwa::protobuf::kn::NoteArchive::decode(&*msg.data)
                {
                    // The note contains a reference to a TSWP.StorageArchive
                    let storage_id = note_archive.contained_storage.identifier;
                    if let Some(storage_obj) =
                        self.object_index.resolve_object(&self.bundle, storage_id)?
                    {
                        for storage_msg in &storage_obj.messages {
                            if let Ok(storage) = crate::iwa::protobuf::tswp::StorageArchive::decode(
                                &*storage_msg.data,
                            ) {
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

    /// Extract presentation metadata.
    ///
    /// Returns metadata from the Keynote bundle's Properties.plist file.
    /// This includes document properties like title, author, creation date, etc.
    ///
    /// # Performance
    ///
    /// This method performs minimal parsing, extracting only standard metadata
    /// fields from the bundle's Properties.plist. The metadata is not cached
    /// within KeynoteDocument to avoid duplication with the Presentation cache.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::keynote::KeynoteDocument;
    ///
    /// let doc = KeynoteDocument::open("presentation.key")?;
    /// if let Some(metadata) = doc.metadata()? {
    ///     if let Some(title) = metadata.title {
    ///         println!("Title: {}", title);
    ///     }
    ///     if let Some(author) = metadata.author {
    ///         println!("Author: {}", author);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    #[allow(unused_assignments)] // has_data is intentionally reassigned to track if any field was set
    pub fn metadata(&self) -> Result<Option<crate::common::Metadata>> {
        let bundle_metadata = self.bundle.metadata();

        // Extract standard metadata fields from Properties.plist and bundle structure
        let mut metadata = crate::common::Metadata::default();
        let mut has_data = false;

        // Extract title (Keynote may store in show structure, try there first)
        let show_title = self.show().ok().and_then(|show| show.title);
        if let Some(title) = show_title {
            metadata.title = Some(title);
            has_data = true;
        }

        // Try alternative title keys from Properties.plist
        if metadata.title.is_none() {
            if let Some(title) = bundle_metadata.get_property_string("Title") {
                metadata.title = Some(title);
                has_data = true;
            } else if let Some(title) = bundle_metadata.get_property_string("kDocumentTitleKey") {
                metadata.title = Some(title);
                has_data = true;
            }
        }

        // Extract author
        if let Some(author) = bundle_metadata.get_property_string("Author") {
            metadata.author = Some(author);
            has_data = true;
        } else if let Some(author) = bundle_metadata.get_property_string("kDocumentAuthorKey") {
            metadata.author = Some(author);
            has_data = true;
        } else if let Some(author) = bundle_metadata.get_property_string("kSFWPAuthorPropertyKey") {
            metadata.author = Some(author);
            has_data = true;
        }

        // Extract keywords
        if let Some(keywords) = bundle_metadata.get_property_string("Keywords") {
            metadata.keywords = Some(keywords);
            has_data = true;
        }

        // Extract comments/description
        if let Some(comments) = bundle_metadata.get_property_string("Comments") {
            metadata.description = Some(comments);
            has_data = true;
        }

        // Extract application name (Keynote applications)
        if let Some(app) = bundle_metadata.detected_application.as_ref() {
            metadata.application = Some(app.clone());
            has_data = true;
        } else {
            // Default to Keynote if not detected
            metadata.application = Some("Keynote".to_string());
            has_data = true;
        }

        // Extract revision from Properties.plist
        if let Some(revision) = bundle_metadata.get_property_string("revision") {
            metadata.revision = Some(revision);
            has_data = true;
        }

        // Extract build version as additional version info
        if let Some(version) = bundle_metadata.latest_build_version() {
            // If we don't have revision yet, use build version
            if metadata.revision.is_none() {
                metadata.revision = Some(version.to_string());
                has_data = true;
            }
        }

        // Extract file format version
        if let Some(format_version) = bundle_metadata.get_property_string("fileFormatVersion") {
            // Store in content_status as it doesn't have a perfect mapping
            metadata.content_status = Some(format!("Keynote Format Version {}", format_version));
            has_data = true;
        }

        // Note: User-facing metadata like creation date, modification date, etc.
        // are typically stored in DocumentMetadata.iwa or Metadata.iwa files,
        // which would require additional IWA parsing. The current implementation
        // extracts what's readily available from Properties.plist and show structure.

        // If we found any metadata, return it
        if has_data {
            Ok(Some(metadata))
        } else {
            Ok(None)
        }
    }

    /// Extract the full show structure with all slides
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::keynote::KeynoteDocument;
    ///
    /// let doc = KeynoteDocument::open("presentation.key")?;
    /// let show = doc.show()?;
    ///
    /// println!("Presentation: {}", show.title.unwrap_or_default());
    /// println!("Slides: {}", show.slide_count());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn show(&self) -> Result<KeynoteShow> {
        let mut show = KeynoteShow::new();

        // Extract show metadata from ShowArchive (message type 2 is KN.ShowArchive)
        let show_objects = self.bundle.find_objects_by_type(1101);
        if let Some((_archive_name, object)) = show_objects.first() {
            let text_parts = object.extract_text();
            show.title = text_parts.first().cloned();
        }

        // Add all slides
        let slides = self.slides()?;
        for slide in slides {
            show.add_slide(slide);
        }

        Ok(show)
    }

    /// Get the underlying bundle
    pub fn bundle(&self) -> &Bundle {
        &self.bundle
    }

    /// Get the object index
    pub fn object_index(&self) -> &ObjectIndex {
        &self.object_index
    }

    /// Get document statistics
    pub fn stats(&self) -> KeynoteDocumentStats {
        let total_objects = self.object_index.all_object_ids().len();
        let slides_result = self.slides();
        let slide_count = slides_result.as_ref().map(|s| s.len()).unwrap_or(0);

        KeynoteDocumentStats {
            total_objects,
            slide_count,
            application: Application::Keynote,
        }
    }
}

/// Statistics about a Keynote document
#[derive(Debug, Clone)]
pub struct KeynoteDocumentStats {
    /// Total number of objects
    pub total_objects: usize,
    /// Number of slides
    pub slide_count: usize,
    /// Application type (always Keynote)
    pub application: Application,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_keynote_document_open() {
        let doc_path = std::path::Path::new("test.key");
        if !doc_path.exists() {
            // Skip test if test file doesn't exist
            return;
        }

        let doc_result = KeynoteDocument::open(doc_path);
        assert!(
            doc_result.is_ok(),
            "Failed to open Keynote document: {:?}",
            doc_result.err()
        );

        let doc = doc_result.unwrap();
        assert!(!doc.object_index.all_object_ids().is_empty());
    }

    #[test]
    fn test_keynote_text_extraction() {
        let doc_path = std::path::Path::new("test.key");
        if !doc_path.exists() {
            return;
        }

        let doc = KeynoteDocument::open(doc_path).unwrap();
        let text_result = doc.text();
        assert!(text_result.is_ok());
    }

    #[test]
    fn test_keynote_slides() {
        let doc_path = std::path::Path::new("test.key");
        if !doc_path.exists() {
            return;
        }

        let doc = KeynoteDocument::open(doc_path).unwrap();
        let slides_result = doc.slides();
        assert!(slides_result.is_ok());

        let slides = slides_result.unwrap();
        // Presentation should have at least one slide
        assert!(
            !slides.is_empty(),
            "Presentation should have at least one slide"
        );
    }

    #[test]
    fn test_keynote_show() {
        let doc_path = std::path::Path::new("test.key");
        if !doc_path.exists() {
            return;
        }

        let doc = KeynoteDocument::open(doc_path).unwrap();
        let show_result = doc.show();
        assert!(show_result.is_ok());

        let show = show_result.unwrap();
        assert!(!show.is_empty(), "Show should have slides");
    }
}
