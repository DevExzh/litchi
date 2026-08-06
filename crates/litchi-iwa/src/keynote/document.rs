//! Keynote Document Implementation
//!
//! Provides high-level API for working with Apple Keynote presentations.

use std::path::Path;
use std::sync::{Arc, OnceLock};

use crate::application::Application;
use crate::bundle::{Bundle, BundleLimits};
use crate::detect::detect_application_from_document;
use crate::object_index::ObjectIndex;
use crate::text::TextExtractor;
use crate::{Error, Result};
use litchi_keynote::{AnimationType, Build, Document, Effect, Seconds, Show, Slide, Transition};
use litchi_keynote::{Mode, Settings, Size};

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
    /// Archive-free semantic state initialized on first semantic access.
    semantic_document: OnceLock<Document>,
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
                semantic_document: OnceLock::new(),
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

    /// Decode slides from the archive while building the semantic snapshot.
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
    ///     println!("Slide {}", slide.index() + 1);
    ///     if let Some(title) = slide.title() {
    ///         println!("  Title: {}", title);
    ///     }
    ///     for text in slide.text_content() {
    ///         println!("  - {}", text);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    fn decode_slides(&self) -> Result<Vec<Slide>> {
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

    /// Borrow the detached semantic slides without reparsing the archive.
    pub fn slides(&self) -> Result<&[Slide]> {
        Ok(self.semantic_document()?.slides())
    }

    fn semantic_document(&self) -> Result<&Document> {
        if let Some(document) = self.state.semantic_document.get() {
            return Ok(document);
        }

        let document = Document::from_show(self.decode_show()?);
        let _ = self.state.semantic_document.set(document);
        self.state.semantic_document.get().ok_or_else(|| {
            Error::ParseError("Keynote semantic state is not initialized".to_owned())
        })
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
    fn parse_slide(&self, index: usize, object: &crate::archive::ArchiveObject) -> Result<Slide> {
        use prost::Message;

        let mut builder = Slide::builder(index);

        // Extract text content from the slide object
        let text_parts = crate::archive::extract_text(object);

        if !text_parts.is_empty() {
            // First text part is typically the title or slide name
            builder.set_title(text_parts.first().cloned());

            // Remaining parts are content
            for text in text_parts.into_iter().skip(1) {
                builder.push_text(text);
            }
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
                    builder.set_title(Some(name.clone()));
                }

                // Extract build animations
                for build_ref in &slide_archive.builds {
                    builder.push_build(self.extract_build_animation(build_ref.identifier)?);
                }

                // Extract transition
                builder.set_transition(Some(self.parse_transition(&slide_archive.transition)?));

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
                        builder.set_title(Some(title));
                    }
                }
                if let Some(identifier) = body_placeholder {
                    let body = self.extract_drawable_text(identifier)?;
                    if !body.is_empty() {
                        builder.push_text(body);
                    }
                }

                // Resolve other drawable references to get text boxes.
                for drawable_ref in &slide_archive.owned_drawables {
                    if Some(drawable_ref.identifier) == title_placeholder
                        || Some(drawable_ref.identifier) == body_placeholder
                    {
                        continue;
                    }
                    let text_content = self.extract_drawable_text(drawable_ref.identifier)?;
                    if !text_content.is_empty() {
                        builder.push_text(text_content);
                    }
                }

                // Extract speaker notes
                if let Some(ref note_ref) = slide_archive.note {
                    let notes = self.extract_speaker_notes(note_ref.identifier)?;
                    if !notes.is_empty() {
                        builder.set_notes(Some(notes));
                    }
                }
            }
        }

        // Extract text from text storages
        let extractor = TextExtractor::new();
        let storage = extractor.extract_from_object(object)?;
        if !storage.is_empty() {
            builder.push_text_storage(storage);
        }

        Ok(builder.build())
    }

    /// Extract build animation from a BuildArchive object
    fn extract_build_animation(&self, build_id: u64) -> Result<Build> {
        use prost::Message;

        if let Some(resolved) = self
            .state
            .object_index
            .resolve_ref_id(&self.state.bundle, build_id)?
        {
            for msg in resolved.messages {
                if let Ok(build_archive) = crate::protobuf::kn::BuildArchive::decode(&*msg.data) {
                    let animation_type = AnimationType::from_identifier(&build_archive.delivery)
                        .map_err(|error| {
                            Error::ParseError(format!("invalid Keynote build identifier: {error}"))
                        })?;
                    let duration = build_archive
                        .attributes
                        .animation_attributes
                        .as_ref()
                        .and_then(|attributes| attributes.duration)
                        .or_else(|| Self::legacy_build_duration(&build_archive))
                        .unwrap_or(0.0);
                    let duration = Seconds::new(duration).map_err(|error| {
                        Error::ParseError(format!("invalid Keynote build duration: {error}"))
                    })?;

                    return Ok(Build::new(animation_type, duration));
                }
            }
        }

        Err(Error::ParseError(format!(
            "Keynote build object {build_id} has no BuildArchive payload"
        )))
    }

    /// Parse transition archive into slide transition
    fn parse_transition(
        &self,
        transition: &crate::protobuf::kn::TransitionArchive,
    ) -> Result<Transition> {
        // Extract duration from attributes
        // The attributes field is required (not Optional)
        let duration = transition
            .attributes
            .animation_attributes
            .as_ref()
            .and_then(|attributes| attributes.duration)
            .or_else(|| Self::legacy_transition_duration(&transition.attributes))
            .unwrap_or(0.0);
        let duration = Seconds::new(duration).map_err(|error| {
            Error::ParseError(format!("invalid Keynote transition duration: {error}"))
        })?;

        Ok(Transition::new(
            transition_effect(&transition.attributes)?,
            duration,
        ))
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

        let resolved = self
            .state
            .object_index
            .resolve_ref_id(&self.state.bundle, drawable_id)?
            .ok_or_else(|| {
                Error::ParseError(format!("Keynote drawable object {drawable_id} is missing"))
            })?;

        let mut storage_id = None;
        for msg in resolved.messages {
            if let Ok(placeholder) =
                crate::protobuf::kn::PlaceholderArchive::decode(msg.data.as_slice())
                && let Some(reference) = placeholder.super_.owned_storage
            {
                storage_id = Some(reference.identifier);
                break;
            }
            if let Ok(shape) = crate::protobuf::tswp::ShapeInfoArchive::decode(msg.data.as_slice())
                && let Some(reference) = shape.owned_storage
            {
                storage_id = Some(reference.identifier);
                break;
            }
        }

        let Some(storage_id) = storage_id else {
            // Images and other unsupported drawables remain preserved in the
            // archive but do not contribute semantic text.
            return Ok(String::new());
        };
        let storage_object = self
            .state
            .object_index
            .resolve_ref_id(&self.state.bundle, storage_id)?
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Keynote drawable storage object {storage_id} is missing"
                ))
            })?;
        for message in storage_object.messages {
            if let Ok(storage) =
                crate::protobuf::tswp::StorageArchive::decode(message.data.as_slice())
            {
                return Ok(storage.text.concat());
            }
        }

        Err(Error::ParseError(format!(
            "Keynote drawable storage object {storage_id} has no StorageArchive payload"
        )))
    }

    /// Extract speaker notes from a NoteArchive object
    fn extract_speaker_notes(&self, note_id: u64) -> Result<String> {
        use prost::Message;

        let resolved = self
            .state
            .object_index
            .resolve_ref_id(&self.state.bundle, note_id)?
            .ok_or_else(|| {
                Error::ParseError(format!("Keynote notes object {note_id} is missing"))
            })?;

        for msg in resolved.messages {
            if let Ok(note_archive) = crate::protobuf::kn::NoteArchive::decode(&*msg.data) {
                // The note contains a reference to a TSWP.StorageArchive.
                let storage_id = note_archive.contained_storage.identifier;
                let storage_obj = self
                    .state
                    .object_index
                    .resolve_ref_id(&self.state.bundle, storage_id)?
                    .ok_or_else(|| {
                        Error::ParseError(format!(
                            "Keynote notes storage object {storage_id} is missing"
                        ))
                    })?;
                for storage_msg in storage_obj.messages {
                    if let Ok(storage) =
                        crate::protobuf::tswp::StorageArchive::decode(&*storage_msg.data)
                    {
                        return Ok(storage.text.join("\n"));
                    }
                }
            }
        }

        Err(Error::ParseError(format!(
            "Keynote notes object {note_id} has no NoteArchive payload"
        )))
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
    /// # Errors
    ///
    /// Returns an error when the Keynote show cannot be decoded.
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
        let show_title = self.show()?.title().map(str::to_owned);
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
        if let Some(app) = bundle_metadata.detected_application() {
            metadata.application = Some(app.to_owned());
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

    fn decode_show(&self) -> Result<Show> {
        use prost::Message;

        let mut builder = Show::builder();

        let document = Self::root_document(&self.state.bundle)?;
        let show_object = self
            .bundle_object(document.show.identifier)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Keynote show object {} is missing",
                    document.show.identifier
                ))
            })?;
        let show_archive = show_object
            .messages
            .iter()
            .find_map(|message| {
                crate::protobuf::kn::ShowArchive::decode(message.data.as_slice()).ok()
            })
            .ok_or_else(|| Error::ParseError("Keynote show payload is missing".to_owned()))?;
        builder.set_settings(settings_from_show_archive(&show_archive)?);
        let text_parts = crate::archive::extract_text(show_object);
        builder.set_title(text_parts.first().cloned());

        for slide in self.decode_slides()? {
            builder.push_slide(slide);
        }

        Ok(builder.build())
    }

    /// Extract the full detached show structure with all slides without
    /// reparsing the archive.
    ///
    /// # Errors
    ///
    /// Returns an error when the root document, show archive, settings, or a
    /// referenced slide cannot be decoded.
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
    /// println!("Presentation: {}", show.title().unwrap_or_default());
    /// println!("Slides: {}", show.slide_count());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// # }
    /// ```
    pub fn show(&self) -> Result<&Show> {
        Ok(self.semantic_document()?.show())
    }

    /// Return a cheap handle to the archive-free semantic snapshot built on demand.
    ///
    /// # Errors
    ///
    /// Returns an error when the underlying show or one of its referenced
    /// slides cannot be decoded.
    pub fn semantic_snapshot(&self) -> Result<Document> {
        Ok(self.semantic_document()?.snapshot())
    }

    /// Validate this immutable snapshot without exposing archive internals.
    ///
    /// Validation is performed directly against the private parsed bundle and
    /// returns the typed crate result. Callers can validate a document without
    /// depending on archive or object-index representations.
    pub fn validate(&self) -> Result<()> {
        self.state.bundle.validate()
    }

    /// Get document statistics after resolving the presentation slides.
    pub fn stats(&self) -> Result<KeynoteDocumentStats> {
        let total_objects = self.state.object_index.object_count();
        let slide_count = self.slides()?.len();

        Ok(KeynoteDocumentStats {
            total_objects,
            slide_count,
            application: Application::Keynote,
        })
    }
}

#[allow(deprecated)]
fn transition_effect(
    attributes: &crate::protobuf::kn::TransitionAttributesArchive,
) -> Result<Effect> {
    let Some(identifier) = attributes
        .animation_attributes
        .as_ref()
        .and_then(|animation| animation.effect.as_deref())
        .or(attributes.database_effect.as_deref())
    else {
        return Ok(Effect::None);
    };
    Effect::from_identifier(identifier)
        .map_err(|error| Error::ParseError(format!("invalid Keynote transition effect: {error}")))
}

fn settings_from_show_archive(show: &crate::protobuf::kn::ShowArchive) -> Result<Settings> {
    let size = Size::new(show.size.width, show.size.height)
        .map_err(|error| Error::ParseError(format!("invalid Keynote show size: {error}")))?;
    let mut settings = Settings::new(size);
    settings.set_slide_numbers_visible(show.slide_numbers_visible);
    settings.set_loop_presentation(show.loop_presentation);
    settings
        .set_mode(show.mode.map(Mode::from_raw))
        .map_err(|error| Error::ParseError(format!("invalid Keynote show mode: {error}")))?;
    settings.set_autoplay_transition_delay(
        show.autoplay_transition_delay
            .map(Seconds::new)
            .transpose()
            .map_err(|error| {
                Error::ParseError(format!("invalid Keynote transition delay: {error}"))
            })?,
    );
    settings.set_autoplay_build_delay(
        show.autoplay_build_delay
            .map(Seconds::new)
            .transpose()
            .map_err(|error| Error::ParseError(format!("invalid Keynote build delay: {error}")))?,
    );
    settings.set_idle_timer_active(show.idle_timer_active);
    settings.set_idle_timer_delay(
        show.idle_timer_delay
            .map(Seconds::new)
            .transpose()
            .map_err(|error| Error::ParseError(format!("invalid Keynote idle delay: {error}")))?,
    );
    settings.set_automatically_plays_upon_open(show.automatically_plays_upon_open);
    settings
        .validate()
        .map_err(|error| Error::ParseError(format!("invalid Keynote show settings: {error}")))?;
    Ok(settings)
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
#[allow(deprecated)]
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
        assert!(doc.stats().unwrap().total_objects > 0);
        assert!(doc.validate().is_ok());
    }

    #[test]
    fn keynote_from_bytes_with_limits_enforces_input_budget() {
        let limits = BundleLimits::new(1, 10, 100, 100, 100).unwrap();
        let error = KeynoteDocument::from_bytes_with_limits(&[0, 1], limits).unwrap_err();
        assert!(error.to_string().contains("iWork bundle input"));
    }

    #[test]
    fn transition_effect_prefers_modern_identifier_and_preserves_unknown_values() {
        let attributes = crate::protobuf::kn::TransitionAttributesArchive {
            animation_attributes: Some(crate::protobuf::kn::AnimationAttributesArchive {
                effect: Some("com.example.future-transition".to_owned()),
                ..Default::default()
            }),
            database_effect: Some("apple:dissolve".to_owned()),
            ..Default::default()
        };
        assert_eq!(
            transition_effect(&attributes).unwrap(),
            Effect::Unknown {
                identifier: "com.example.future-transition".to_owned().into_boxed_str()
            }
        );
    }

    #[test]
    fn transition_effect_falls_back_to_legacy_identifier_and_defaults_to_none() {
        let legacy = crate::protobuf::kn::TransitionAttributesArchive {
            database_effect: Some("apple:dissolve".to_owned()),
            ..Default::default()
        };
        assert_eq!(transition_effect(&legacy).unwrap(), Effect::Dissolve);
        assert_eq!(
            transition_effect(&Default::default()).unwrap(),
            Effect::None
        );
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
