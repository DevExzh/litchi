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
        // - transition: reference to KN.TransitionArchive
        // - master_slide: reference to master slide

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

                // TODO: Extract build animations
                // The builds field contains references to KN.BuildArchive objects
                // which define the animation effects for objects on the slide

                // TODO: Extract transition
                // The transition field contains a reference to a transition effect

                // TODO: Resolve drawable references to get text boxes and other content
                // The drawables field contains references to TSD.DrawableArchive objects
                // which can include text boxes, shapes, images, etc.

                // TODO: Extract speaker notes
                // The note field contains a reference to KN.NoteArchive
                // which has the speaker notes text
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
        assert!(doc.object_index.all_object_ids().len() > 0);
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
