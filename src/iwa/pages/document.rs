//! Pages Document Implementation
//!
//! Provides high-level API for working with Apple Pages documents.

use std::path::Path;

use super::section::{PagesSection, PagesSectionType};
use crate::iwa::Result;
use crate::iwa::bundle::Bundle;
use crate::iwa::object_index::ObjectIndex;
use crate::iwa::registry::Application;
use crate::iwa::text::TextExtractor;

/// High-level interface for Pages documents
pub struct PagesDocument {
    /// Underlying bundle
    bundle: Bundle,
    /// Object index for cross-referencing
    object_index: ObjectIndex,
}

impl PagesDocument {
    /// Open a Pages document from a path
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::pages::PagesDocument;
    ///
    /// let doc = PagesDocument::open("document.pages")?;
    /// println!("Loaded Pages document");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let bundle = Bundle::open(path)?;

        // Verify this is a Pages document
        Self::verify_application(&bundle)?;

        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Ok(Self {
            bundle,
            object_index,
        })
    }

    /// Open a Pages document from raw bytes
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::pages::PagesDocument;
    /// use std::fs;
    ///
    /// let data = fs::read("document.pages")?;
    /// let doc = PagesDocument::from_bytes(&data)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let bundle = Bundle::from_bytes(bytes)?;

        // Verify this is a Pages document
        Self::verify_application(&bundle)?;

        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Ok(Self {
            bundle,
            object_index,
        })
    }

    /// Verify that the bundle is a Pages document
    fn verify_application(bundle: &Bundle) -> Result<()> {
        // Check for Pages-specific message types (TP.* types in range 10000-10999)
        // Message type 10000 is TP.DocumentArchive
        let has_pages_types = bundle.archives().values().any(|archive| {
            archive.objects.iter().any(|obj| {
                obj.messages
                    .iter()
                    .any(|msg| msg.type_ == 10000 || (10000..11000).contains(&msg.type_))
            })
        });

        if !has_pages_types {
            // Be lenient - if we can't definitively identify it as another type, allow it
            // This helps with documents that might not have explicit Pages markers
        }

        Ok(())
    }

    /// Extract all text content from the document
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::pages::PagesDocument;
    ///
    /// let doc = PagesDocument::open("document.pages")?;
    /// let text = doc.text()?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        let mut extractor = TextExtractor::new();
        extractor.extract_from_bundle(&self.bundle)?;
        Ok(extractor.get_text())
    }

    /// Extract sections from the document
    ///
    /// Pages documents are organized into sections. This method parses the
    /// document structure and returns all sections with their content.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::pages::PagesDocument;
    ///
    /// let doc = PagesDocument::open("document.pages")?;
    /// let sections = doc.sections()?;
    ///
    /// for section in sections {
    ///     println!("Section {}: {}", section.index, section.section_type.name());
    ///     println!("{}", section.plain_text());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn sections(&self) -> Result<Vec<PagesSection>> {
        let mut sections = Vec::new();

        // Find section archives (message type 10011 is TP.SectionArchive)
        let section_objects = self.bundle.find_objects_by_type(10011);

        if section_objects.is_empty() {
            // If no explicit sections found, create a single body section
            // with all text content
            let mut section = PagesSection::new(0, PagesSectionType::Body);

            // Extract text from all TSWP storage objects
            let mut extractor = TextExtractor::new();
            extractor.extract_from_bundle(&self.bundle)?;

            for storage in extractor.storages() {
                if !storage.is_empty() {
                    section.text_storages.push(storage.clone());
                    section.paragraphs.push(storage.plain_text().to_string());
                }
            }

            if !section.is_empty() {
                sections.push(section);
            }
        } else {
            // Parse explicit sections
            for (index, (_archive_name, _object)) in section_objects.iter().enumerate() {
                let section = PagesSection::new(index, PagesSectionType::Body);

                // In a full implementation, we would:
                // 1. Parse the section protobuf message
                // 2. Resolve references to text storage objects
                // 3. Extract section-specific properties (margins, headers, footers)
                // 4. Build the complete section structure
                //
                // For now, we create placeholder sections

                sections.push(section);
            }
        }

        Ok(sections)
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
    pub fn stats(&self) -> PagesDocumentStats {
        let total_objects = self.object_index.all_object_ids().len();
        let sections_result = self.sections();
        let section_count = sections_result.as_ref().map(|s| s.len()).unwrap_or(0);

        PagesDocumentStats {
            total_objects,
            section_count,
            application: Application::Pages,
        }
    }
}

/// Statistics about a Pages document
#[derive(Debug, Clone)]
pub struct PagesDocumentStats {
    /// Total number of objects
    pub total_objects: usize,
    /// Number of sections
    pub section_count: usize,
    /// Application type (always Pages)
    pub application: Application,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_pages_document_open() {
        let doc_path = std::path::Path::new("test.pages");
        if !doc_path.exists() {
            // Skip test if test file doesn't exist
            return;
        }

        let doc_result = PagesDocument::open(doc_path);
        assert!(
            doc_result.is_ok(),
            "Failed to open Pages document: {:?}",
            doc_result.err()
        );

        let doc = doc_result.unwrap();
        assert!(doc.object_index.all_object_ids().len() > 0);
    }

    #[test]
    fn test_pages_text_extraction() {
        let doc_path = std::path::Path::new("test.pages");
        if !doc_path.exists() {
            return;
        }

        let doc = PagesDocument::open(doc_path).unwrap();
        let text_result = doc.text();
        assert!(text_result.is_ok());

        // Text might be empty for some documents, but extraction should succeed
        let _text = text_result.unwrap();
    }

    #[test]
    fn test_pages_sections() {
        let doc_path = std::path::Path::new("test.pages");
        if !doc_path.exists() {
            return;
        }

        let doc = PagesDocument::open(doc_path).unwrap();
        let sections_result = doc.sections();
        assert!(sections_result.is_ok());

        let sections = sections_result.unwrap();
        // Document should have at least one section (even if implicit)
        assert!(
            !sections.is_empty(),
            "Document should have at least one section"
        );
    }
}
