//! Keynote Document Implementation
//!
//! Provides high-level API for working with Apple Keynote presentations.

use std::path::Path;
use std::sync::Arc;

use super::show::KeynoteShow;
use super::slide::KeynoteSlide;
use crate::bundle::{Bundle, BundleLimits};
use crate::object_index::ObjectIndex;
use crate::registry::{Application, detect_application_from_document};
use crate::text::TextExtractor;
use crate::{Error, Result};

/// High-level interface for Keynote documents
#[derive(Debug, Clone)]
pub struct KeynoteDocument {
    state: Arc<KeynoteDocumentState>,
}

#[derive(Debug)]
struct KeynoteDocumentState {
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
    /// use litchi_iwa::keynote::KeynoteDocument;
    ///
    /// let doc = KeynoteDocument::open("presentation.key")?;
    /// println!("Loaded Keynote presentation");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_limits(path, BundleLimits::default())
    }

    /// Open a Keynote document under caller-selected bundle ingress ceilings.
    pub fn open_with_limits<P: AsRef<Path>>(path: P, limits: BundleLimits) -> Result<Self> {
        let bundle = Bundle::open_with_limits(path, limits)?;
        Self::verify_application(&bundle)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Ok(Self::from_parts(bundle, object_index))
    }

    /// Open a Keynote document from raw bytes
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_iwa::keynote::KeynoteDocument;
    /// use std::fs;
    ///
    /// let data = fs::read("presentation.key")?;
    /// let doc = KeynoteDocument::from_bytes(&data)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, BundleLimits::default())
    }

    /// Open a Keynote document from bytes under caller-selected ingress
    /// ceilings.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: BundleLimits) -> Result<Self> {
        let bundle = Bundle::from_bytes_with_limits(bytes, limits)?;
        Self::verify_application(&bundle)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Ok(Self::from_parts(bundle, object_index))
    }

    fn from_parts(bundle: Bundle, object_index: ObjectIndex) -> Self {
        Self {
            state: Arc::new(KeynoteDocumentState {
                bundle,
                object_index,
            }),
        }
    }

    /// Capture a cheap immutable snapshot that shares all parsed document state.
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Create a Keynote document from raw bytes (ZIP archive data).
    ///
    /// This convenience entry point currently performs the same parsing as
    /// [`Self::from_bytes`]; it does not accept a previously parsed archive.
    pub fn from_archive_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes)
    }

    /// Create a Keynote document from archive bytes under caller-selected
    /// ingress ceilings.
    pub fn from_archive_bytes_with_limits(bytes: &[u8], limits: BundleLimits) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, limits)
    }

    fn verify_application(bundle: &Bundle) -> Result<()> {
        Self::root_document(bundle).map(|_| ())
    }

    fn root_document(bundle: &Bundle) -> Result<crate::protobuf::kn::DocumentArchive> {
        use prost::Message;

        let object = bundle
            .get_archive("Index/Document.iwa")
            .and_then(|archive| archive.object(1))
            .ok_or_else(|| Error::InvalidFormat("Keynote root object 1 is missing".to_owned()))?;
        object
            .messages
            .iter()
            .find(|message| {
                detect_application_from_document(&message.data) == Some(Application::Keynote)
            })
            .and_then(|message| {
                crate::protobuf::kn::DocumentArchive::decode(message.data.as_slice()).ok()
            })
            .ok_or_else(|| {
                Error::InvalidFormat("package does not contain a Keynote root document".to_owned())
            })
    }

    /// Extract all text content from the presentation
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_iwa::keynote::KeynoteDocument;
    ///
    /// let doc = KeynoteDocument::open("presentation.key")?;
    /// let text = doc.text()?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        let mut extractor = TextExtractor::new();
        extractor.extract_from_bundle(&self.state.bundle)?;
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
    /// use litchi_iwa::keynote::KeynoteDocument;
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

        for (index, slide_id) in self.slide_ids()?.into_iter().enumerate() {
            let object = self.bundle_object(slide_id).ok_or_else(|| {
                crate::Error::ParseError(format!(
                    "Keynote slide tree references missing object {slide_id}"
                ))
            })?;
            slides.push(self.parse_slide(index, object)?);
        }

        Ok(slides)
    }

    fn slide_ids(&self) -> Result<Vec<u64>> {
        use prost::Message;

        let document = Self::root_document(&self.state.bundle)?;
        let show_object = self
            .bundle_object(document.show.identifier)
            .ok_or_else(|| {
                crate::Error::ParseError(format!(
                    "Keynote show object {} is missing",
                    document.show.identifier
                ))
            })?;
        let show = show_object
            .messages
            .iter()
            .find_map(|message| {
                crate::protobuf::kn::ShowArchive::decode(message.data.as_slice()).ok()
            })
            .ok_or_else(|| {
                crate::Error::ParseError("Keynote show payload is missing".to_string())
            })?;

        let mut slide_ids = Vec::with_capacity(show.slide_tree.slides.len());
        for node_reference in show.slide_tree.slides {
            let node_object = self
                .bundle_object(node_reference.identifier)
                .ok_or_else(|| {
                    crate::Error::ParseError(format!(
                        "Keynote slide node {} is missing",
                        node_reference.identifier
                    ))
                })?;
            let node = node_object
                .messages
                .iter()
                .find_map(|message| {
                    crate::protobuf::kn::SlideNodeArchive::decode(message.data.as_slice()).ok()
                })
                .ok_or_else(|| {
                    crate::Error::ParseError(format!(
                        "Object {} has no KN.SlideNodeArchive payload",
                        node_reference.identifier
                    ))
                })?;
            if let Some(slide) = node.slide {
                slide_ids.push(slide.identifier);
            }
        }
        Ok(slide_ids)
    }

    fn bundle_object(&self, identifier: u64) -> Option<&crate::archive::ArchiveObject> {
        self.state
            .bundle
            .iter_archives()
            .map(|(_, archive)| archive)
            .find_map(|archive| archive.object(identifier))
    }

    /// Parse a single slide from an object
    fn parse_slide(
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
                if let Some(ref master) = slide_archive.template_slide {
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

                let title_placeholder = slide_archive
                    .title_placeholder
                    .as_ref()
                    .map(|reference| reference.identifier);
                let body_placeholder = slide_archive
                    .body_placeholder
                    .as_ref()
                    .map(|reference| reference.identifier);

                if let Some(identifier) = title_placeholder {
                    let title = self.extract_drawable_text(identifier)?;
                    if !title.is_empty() {
                        slide.title = Some(title);
                    }
                }
                if let Some(identifier) = body_placeholder {
                    let body = self.extract_drawable_text(identifier)?;
                    if !body.is_empty() {
                        slide.text_content.push(body);
                    }
                }

                // Resolve other drawable references to get text boxes.
                for drawable_ref in &slide_archive.owned_drawables {
                    if Some(drawable_ref.identifier) == title_placeholder
                        || Some(drawable_ref.identifier) == body_placeholder
                    {
                        continue;
                    }
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

        if let Some(resolved) = self
            .state
            .object_index
            .resolve_id(&self.state.bundle, build_id)?
        {
            for msg in &resolved.messages {
                if let Ok(build_archive) = crate::protobuf::kn::BuildArchive::decode(&*msg.data) {
                    let animation_type = Self::parse_build_delivery(&build_archive.delivery);
                    let target_id = build_archive.drawable.as_ref().map(|r| r.identifier);
                    let duration = build_archive
                        .attributes
                        .animation_attributes
                        .as_ref()
                        .and_then(|attributes| attributes.duration)
                        .or_else(|| Self::legacy_build_duration(&build_archive))
                        .unwrap_or(0.0) as f32;

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
        transition: &crate::protobuf::kn::TransitionArchive,
    ) -> Option<super::slide::SlideTransition> {
        use super::slide::{SlideTransition, TransitionType};

        // Extract duration from attributes
        // The attributes field is required (not Optional)
        let duration = transition
            .attributes
            .animation_attributes
            .as_ref()
            .and_then(|attributes| attributes.duration)
            .or_else(|| Self::legacy_transition_duration(&transition.attributes))
            .unwrap_or(0.0) as f32;

        // Determine transition type from attributes
        // The actual transition type is embedded in the attributes structure
        // For now, we use a generic transition type
        let transition_type = TransitionType::Other;

        Some(SlideTransition {
            transition_type,
            duration,
        })
    }

    #[allow(deprecated)]
    fn legacy_build_duration(build: &crate::protobuf::kn::BuildArchive) -> Option<f64> {
        build.attributes.database_duration.or(build.duration)
    }

    #[allow(deprecated)]
    fn legacy_transition_duration(
        attributes: &crate::protobuf::kn::TransitionAttributesArchive,
    ) -> Option<f64> {
        attributes.database_duration
    }

    /// Extract text content from a drawable object
    fn extract_drawable_text(&self, drawable_id: u64) -> Result<String> {
        use prost::Message;

        if let Some(resolved) = self
            .state
            .object_index
            .resolve_id(&self.state.bundle, drawable_id)?
        {
            let mut storage_id = None;
            for msg in &resolved.messages {
                if let Ok(placeholder) =
                    crate::protobuf::kn::PlaceholderArchive::decode(msg.data.as_slice())
                    && let Some(reference) = placeholder.super_.owned_storage
                {
                    storage_id = Some(reference.identifier);
                    break;
                }
                if let Ok(shape) =
                    crate::protobuf::tswp::ShapeInfoArchive::decode(msg.data.as_slice())
                    && let Some(reference) = shape.owned_storage
                {
                    storage_id = Some(reference.identifier);
                    break;
                }
            }

            if let Some(storage_id) = storage_id
                && let Some(storage_object) = self
                    .state
                    .object_index
                    .resolve_id(&self.state.bundle, storage_id)?
            {
                for message in storage_object.messages {
                    if let Ok(storage) =
                        crate::protobuf::tswp::StorageArchive::decode(message.data.as_slice())
                    {
                        return Ok(storage.text.concat());
                    }
                }
            }
        }

        Ok(String::new())
    }

    /// Extract speaker notes from a NoteArchive object
    fn extract_speaker_notes(&self, note_id: u64) -> Result<String> {
        use prost::Message;

        if let Some(resolved) = self
            .state
            .object_index
            .resolve_id(&self.state.bundle, note_id)?
        {
            for msg in &resolved.messages {
                if let Ok(note_archive) = crate::protobuf::kn::NoteArchive::decode(&*msg.data) {
                    // The note contains a reference to a TSWP.StorageArchive
                    let storage_id = note_archive.contained_storage.identifier;
                    if let Some(storage_obj) = self
                        .state
                        .object_index
                        .resolve_id(&self.state.bundle, storage_id)?
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
    /// use litchi_iwa::keynote::KeynoteDocument;
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
    pub fn metadata(&self) -> Result<Option<litchi_core::Metadata>> {
        let bundle_metadata = self.state.bundle.metadata();

        // Extract standard metadata fields from Properties.plist and bundle structure
        let mut metadata = litchi_core::Metadata::default();
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
    /// use litchi_iwa::keynote::KeynoteDocument;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let doc = KeynoteDocument::open("presentation.key")?;
    /// let show = doc.show()?;
    ///
    /// println!("Presentation: {}", show.title.as_deref().unwrap_or_default());
    /// println!("Slides: {}", show.slide_count());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// # }
    /// ```
    pub fn show(&self) -> Result<KeynoteShow> {
        let mut show = KeynoteShow::new();

        // Extract show metadata from ShowArchive (message type 2 is KN.ShowArchive)
        let show_objects = self.state.bundle.find_objects_by_type(1101);
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
        &self.state.bundle
    }

    /// Get the object index
    pub fn object_index(&self) -> &ObjectIndex {
        &self.state.object_index
    }

    /// Return a bounded, deterministic validation report for this snapshot.
    pub fn validation_report(&self) -> crate::bundle::BundleValidationReport {
        self.state.bundle.validation_report()
    }

    /// Validate this immutable snapshot without mutating it.
    pub fn validate(&self) -> Result<()> {
        self.validation_report().as_result()
    }

    /// Get document statistics after resolving the presentation slides.
    pub fn stats(&self) -> Result<KeynoteDocumentStats> {
        let total_objects = self.state.object_index.object_ids()?.len();
        let slide_count = self.slides()?.len();

        Ok(KeynoteDocumentStats {
            total_objects,
            slide_count,
            application: Application::Keynote,
        })
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

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn keynote_documents_are_send_and_sync() {
        assert_send_sync::<KeynoteDocument>();
    }

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
        assert!(!doc.object_index().object_ids().unwrap().is_empty());
        assert!(doc.validate().is_ok());
    }

    #[test]
    fn keynote_from_bytes_with_limits_enforces_input_budget() {
        let limits = BundleLimits::new(1, 10, 100, 100, 100).unwrap();
        let error = KeynoteDocument::from_bytes_with_limits(&[0, 1], limits).unwrap_err();
        assert!(error.to_string().contains("iWork bundle input"));
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
