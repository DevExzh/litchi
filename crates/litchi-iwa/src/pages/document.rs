//! Pages Document Implementation
//!
//! Provides high-level API for working with Apple Pages documents.

use std::path::Path;
use std::sync::Arc;

use prost::Message;

use super::section::{PagesSection, PagesSectionType};
use crate::bundle::{Bundle, BundleLimits};
use crate::object_index::ObjectIndex;
use crate::protobuf::{tp, tswp};
use crate::registry::Application;
use crate::text::{TextExtractor, TextStorage};
use crate::{Error, Result};

/// High-level interface for Pages documents
#[derive(Debug, Clone)]
pub struct PagesDocument {
    state: Arc<PagesDocumentState>,
}

#[derive(Debug)]
struct PagesDocumentState {
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
    /// use litchi_iwa::pages::PagesDocument;
    ///
    /// let doc = PagesDocument::open("document.pages")?;
    /// println!("Loaded Pages document");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_limits(path, BundleLimits::default())
    }

    /// Open a Pages document under caller-selected bundle ingress ceilings.
    pub fn open_with_limits<P: AsRef<Path>>(path: P, limits: BundleLimits) -> Result<Self> {
        let bundle = Bundle::open_with_limits(path, limits)?;

        // Verify this is a Pages document
        Self::verify_application(&bundle)?;

        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Ok(Self::from_parts(bundle, object_index))
    }

    /// Open a Pages document from raw bytes
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_iwa::pages::PagesDocument;
    /// use std::fs;
    ///
    /// let data = fs::read("document.pages")?;
    /// let doc = PagesDocument::from_bytes(&data)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, BundleLimits::default())
    }

    /// Open a Pages document from bytes under caller-selected ingress
    /// ceilings.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: BundleLimits) -> Result<Self> {
        let bundle = Bundle::from_bytes_with_limits(bytes, limits)?;

        // Verify this is a Pages document
        Self::verify_application(&bundle)?;

        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Ok(Self::from_parts(bundle, object_index))
    }

    fn from_parts(bundle: Bundle, object_index: ObjectIndex) -> Self {
        Self {
            state: Arc::new(PagesDocumentState {
                bundle,
                object_index,
            }),
        }
    }

    /// Capture a cheap immutable snapshot that shares all parsed document state.
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Create a Pages document from raw bytes (ZIP archive data).
    ///
    /// This convenience entry point currently performs the same parsing as
    /// [`Self::from_bytes`]; it does not accept a previously parsed archive.
    pub fn from_archive_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes)
    }

    /// Create a Pages document from archive bytes under caller-selected
    /// ingress ceilings.
    pub fn from_archive_bytes_with_limits(bytes: &[u8], limits: BundleLimits) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, limits)
    }

    /// Verify that the bundle is a Pages document
    fn verify_application(bundle: &Bundle) -> Result<()> {
        if Self::root_document(bundle).is_none() {
            return Err(Error::InvalidFormat(
                "package does not contain a Pages root document".to_owned(),
            ));
        }
        Ok(())
    }

    fn root_document(bundle: &Bundle) -> Option<tp::DocumentArchive> {
        bundle
            .get_archive("Index/Document.iwa")?
            .object(1)?
            .messages
            .iter()
            .find(|message| message.type_ == 10000)
            .and_then(|message| tp::DocumentArchive::decode(message.data.as_slice()).ok())
    }

    fn body_storage(&self) -> Result<Option<TextStorage>> {
        let Some(reference) =
            Self::root_document(&self.state.bundle).and_then(|doc| doc.body_storage)
        else {
            return Ok(None);
        };
        let Some(object) = self
            .state
            .object_index
            .resolve_id(&self.state.bundle, reference.identifier)?
        else {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage object {} is missing",
                reference.identifier
            )));
        };
        let storage = object
            .messages
            .iter()
            .filter(|message| message.type_ == 2001 || message.type_ == 2022)
            .find_map(|message| tswp::StorageArchive::decode(message.data.as_slice()).ok())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages body object {} has no text storage payload",
                    reference.identifier
                ))
            })?;
        let mut result = TextStorage::from_text(storage.text.concat());
        result.identifier = Some(reference.identifier);
        Ok(Some(result))
    }

    /// Extract all text content from the document
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_iwa::pages::PagesDocument;
    ///
    /// let doc = PagesDocument::open("document.pages")?;
    /// let text = doc.text()?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        let mut extractor = TextExtractor::new();
        extractor.extract_from_bundle(&self.state.bundle)?;
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
    /// use litchi_iwa::pages::PagesDocument;
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
        if let Some(storage) = self.body_storage()? {
            let mut section = PagesSection::new(0, PagesSectionType::Body);
            section.text_storages.push(storage);
            return Ok(vec![section]);
        }

        // Older and page-layout packages may not expose a root body storage.
        // Keep the conservative fallback, but only after verifying the Pages root.
        let mut sections = Vec::new();
        let mut section = PagesSection::new(0, PagesSectionType::Body);
        let mut extractor = TextExtractor::new();
        extractor.extract_from_bundle(&self.state.bundle)?;
        section.text_storages.extend(
            extractor
                .storages()
                .iter()
                .filter(|s| !s.is_empty())
                .cloned(),
        );
        if !section.is_empty() {
            sections.push(section);
        }
        Ok(sections)
    }

    /// Get the underlying bundle
    pub fn bundle(&self) -> &Bundle {
        &self.state.bundle
    }

    /// Return a bounded, deterministic validation report for this snapshot.
    pub fn validation_report(&self) -> crate::bundle::BundleValidationReport {
        self.state.bundle.validation_report()
    }

    /// Validate this immutable snapshot without mutating it.
    pub fn validate(&self) -> Result<()> {
        self.validation_report().as_result()
    }

    /// Get the object index
    pub fn object_index(&self) -> &ObjectIndex {
        &self.state.object_index
    }

    /// Get document statistics after resolving the document sections.
    pub fn stats(&self) -> Result<PagesDocumentStats> {
        let total_objects = self.state.object_index.object_ids()?.len();
        let section_count = self.sections()?.len();

        Ok(PagesDocumentStats {
            total_objects,
            section_count,
            application: Application::Pages,
        })
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
    use crate::IWorkPackage;
    use crate::archive::{Archive, ArchiveObject, RawMessage};
    use crate::protobuf::tsp::Reference;

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn pages_documents_are_send_and_sync() {
        assert_send_sync::<PagesDocument>();
    }

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
        assert!(!doc.object_index().object_ids().unwrap().is_empty());
        assert!(doc.validate().is_ok());
    }

    #[test]
    fn pages_from_bytes_with_limits_enforces_input_budget() {
        let limits = BundleLimits::new(1, 10, 100, 100, 100).unwrap();
        let error = PagesDocument::from_bytes_with_limits(&[0, 1], limits).unwrap_err();
        assert!(error.to_string().contains("iWork bundle input"));
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
    fn resolves_body_storage_from_pages_root() {
        let body_id = 42;
        let root = tp::DocumentArchive {
            body_storage: Some(Reference {
                identifier: body_id,
                ..Default::default()
            }),
            ..Default::default()
        };
        let body = tswp::StorageArchive {
            text: vec!["Pages body — café 東京 🚀".to_owned()],
            ..Default::default()
        };
        let archive = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 10000,
                        data: root.encode_to_vec(),
                    }],
                )
                .unwrap(),
                ArchiveObject::new(
                    body_id,
                    vec![RawMessage {
                        type_: 2001,
                        data: body.encode_to_vec(),
                    }],
                )
                .unwrap(),
            ],
        };
        let mut package = IWorkPackage::new();
        package
            .replace_archive("Index/Document.iwa", &archive)
            .unwrap();

        let document = PagesDocument::from_bytes(&package.to_bytes().unwrap()).unwrap();
        let sections = document.sections().unwrap();
        assert_eq!(sections.len(), 1);
        assert_eq!(sections[0].text_storages[0].identifier, Some(body_id));
        assert_eq!(sections[0].plain_text(), "Pages body — café 東京 🚀");

        let structured =
            crate::structured::extract_sections(document.bundle(), document.object_index())
                .unwrap();
        assert_eq!(structured.len(), 1);
        assert_eq!(structured[0].paragraphs, ["Pages body — café 東京 🚀"]);
    }
}
