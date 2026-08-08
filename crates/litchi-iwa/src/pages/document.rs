//! Pages Document Implementation
//!
//! Provides high-level API for working with Apple Pages documents.

use std::num::NonZeroU64;
use std::path::Path;
use std::sync::Arc;

use prost::Message;

use crate::application::Application;
use crate::bundle::{Bundle, BundleLimits};
use crate::object_index::ObjectIndex;
use crate::protobuf::{tp, tswp};
use crate::text::{Storage, TextExtractor};
use crate::{Error, Result};
use litchi_iwa_index::ObjectId;

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
    /// Immutable Pages semantic snapshot built at ingress.
    document: litchi_pages::Document,
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
    pub(crate) fn open_with_limits<P: AsRef<Path>>(path: P, limits: BundleLimits) -> Result<Self> {
        let bundle = Bundle::open_with_limits(path, limits)?;

        let object_index = ObjectIndex::from_bundle(&bundle)?;
        let root = Self::root_body_reference(&bundle)?;
        let document =
            Self::decode_document(&bundle, &object_index, root, limits.max_iwa_stream_bytes())?;

        Ok(Self::from_parts(bundle, object_index, document))
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
    pub(crate) fn from_bytes_with_limits(bytes: &[u8], limits: BundleLimits) -> Result<Self> {
        let bundle = Bundle::from_bytes_with_limits(bytes, limits)?;

        let object_index = ObjectIndex::from_bundle(&bundle)?;
        let root = Self::root_body_reference(&bundle)?;
        let document =
            Self::decode_document(&bundle, &object_index, root, limits.max_iwa_stream_bytes())?;

        Ok(Self::from_parts(bundle, object_index, document))
    }

    fn from_parts(
        bundle: Bundle,
        object_index: ObjectIndex,
        document: litchi_pages::Document,
    ) -> Self {
        Self {
            state: Arc::new(PagesDocumentState {
                bundle,
                object_index,
                document,
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

    /// Decode the native root reference while keeping protobuf details below
    /// the semantic Pages boundary.
    fn root_body_reference(bundle: &Bundle) -> Result<Option<NonZeroU64>> {
        let archive = bundle.get_archive("Index/Document.iwa").ok_or_else(|| {
            Error::InvalidFormat("package does not contain a Pages root document".to_owned())
        })?;
        let object = archive
            .object(1)
            .ok_or_else(|| Error::InvalidFormat("Pages root object 1 is missing".to_owned()))?;
        let mut payload = None;
        for message in &object.messages {
            if message.type_ == 10_000 && payload.replace(message.data.as_slice()).is_some() {
                return Err(Error::InvalidFormat(
                    "Pages root contains duplicate type-10000 payloads".to_owned(),
                ));
            }
        }
        let payload = payload.ok_or_else(|| {
            Error::InvalidFormat("Pages root has no type-10000 payload".to_owned())
        })?;
        let root = tp::DocumentArchive::decode(payload)?;
        root.body_storage
            .map(|reference| {
                NonZeroU64::new(reference.identifier).ok_or_else(|| {
                    Error::InvalidFormat("Pages root body-storage reference is zero".to_owned())
                })
            })
            .transpose()
    }

    fn decode_document(
        bundle: &Bundle,
        object_index: &ObjectIndex,
        body_identifier: Option<NonZeroU64>,
        max_text_bytes: usize,
    ) -> Result<litchi_pages::Document> {
        let body = match body_identifier {
            Some(identifier) => {
                let object_id = ObjectId::new(identifier.get()).ok_or_else(|| {
                    Error::InvalidFormat("Pages body-storage reference is zero".to_owned())
                })?;
                let object = object_index.resolve(bundle, object_id)?.ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages body storage object {} is missing",
                        identifier
                    ))
                })?;
                let storage =
                    Self::decode_body_storage(&object.messages, identifier, max_text_bytes)?;
                Some(litchi_pages::Body::with_max_text_bytes(
                    vec![storage],
                    max_text_bytes,
                )?)
            },
            None => {
                let mut extractor = TextExtractor::new();
                extractor.extract_from_bundle(bundle)?;
                let storages = extractor.storages().to_vec();
                if storages.is_empty() {
                    None
                } else {
                    Some(litchi_pages::Body::with_max_text_bytes(
                        storages,
                        max_text_bytes,
                    )?)
                }
            },
        };
        let root = body.map_or_else(litchi_pages::Root::empty, litchi_pages::Root::with_body);
        litchi_pages::Document::from_root_with_max_text_bytes(root, max_text_bytes)
            .map_err(Into::into)
    }

    fn decode_body_storage(
        messages: &[crate::archive::RawMessage],
        identifier: NonZeroU64,
        max_text_bytes: usize,
    ) -> Result<Storage> {
        let mut payload = None;
        for message in messages {
            if matches!(message.type_, 2001 | 2022)
                && payload
                    .replace((message.type_, message.data.as_slice()))
                    .is_some()
            {
                return Err(Error::InvalidFormat(format!(
                    "Pages body storage object {} contains duplicate text payloads",
                    identifier
                )));
            }
        }
        let (message_type, payload) = payload.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages body object {} has no type-2001/type-2022 text payload",
                identifier
            ))
        })?;
        let storage = tswp::StorageArchive::decode(payload).map_err(|error| {
            Error::InvalidFormat(format!(
                "Pages body object {} type-{message_type} payload is invalid: {error}",
                identifier
            ))
        })?;
        let text_len = storage.text.iter().try_fold(0usize, |length, line| {
            length.checked_add(line.len()).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages body object {} text length overflows usize",
                    identifier
                ))
            })
        })?;
        if text_len > max_text_bytes {
            return Err(Error::InvalidFormat(format!(
                "Pages body object {} text exceeds {max_text_bytes} bytes",
                identifier
            )));
        }
        let mut text = String::with_capacity(text_len);
        for line in &storage.text {
            text.push_str(line);
        }
        Ok(Storage::from_text(text))
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
        Ok(self.state.document.plain_text())
    }

    /// Borrow the immutable semantic sections in source order.
    ///
    /// Pages documents are organized into sections. The document structure is
    /// decoded at ingress, so this method only borrows the retained semantic
    /// sections and performs no allocation or parsing.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_iwa::pages::PagesDocument;
    ///
    /// let doc = PagesDocument::open("document.pages")?;
    /// let sections = doc.sections();
    ///
    /// for section in sections {
    ///     println!("Section {}: {}", section.index(), section.section_type().name());
    ///     println!("{}", section.plain_text());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    #[must_use]
    pub fn sections(&self) -> &[litchi_pages::Section] {
        self.state.document.sections()
    }

    /// Borrow the immutable Pages semantic snapshot without exposing package
    /// archives, protobuf messages, or native object identifiers.
    #[must_use]
    pub fn semantic_document(&self) -> &litchi_pages::Document {
        &self.state.document
    }

    /// Extract standard document metadata from the Pages package.
    ///
    /// The metadata is parsed while the bundle is opened and projected into
    /// the format-neutral semantic type. Archive entries, plist containers,
    /// and native identifiers remain private to the decoder.
    #[must_use]
    pub fn metadata(&self) -> litchi_core::Metadata {
        let bundle_metadata = self.state.bundle.metadata();
        let revision = bundle_metadata
            .get_property_string("revision")
            .or_else(|| bundle_metadata.latest_build_version().map(str::to_owned));
        let content_status = bundle_metadata
            .get_property_string("fileFormatVersion")
            .map(|version| format!("Pages Format Version {version}"));
        let identifier = bundle_metadata.document_identifier().map(str::to_owned);
        let application = bundle_metadata
            .detected_application()
            .map(str::to_owned)
            .unwrap_or_else(|| "Pages".to_owned());

        litchi_core::Metadata {
            title: bundle_metadata
                .get_property_string("Title")
                .or_else(|| bundle_metadata.get_property_string("kDocumentTitleKey")),
            author: bundle_metadata
                .get_property_string("Author")
                .or_else(|| bundle_metadata.get_property_string("kDocumentAuthorKey"))
                .or_else(|| bundle_metadata.get_property_string("kSFWPAuthorPropertyKey")),
            keywords: bundle_metadata.get_property_string("Keywords"),
            description: bundle_metadata.get_property_string("Comments"),
            application: Some(application),
            revision,
            content_status,
            identifier,
            ..Default::default()
        }
    }

    /// Validate this immutable snapshot without mutating it.
    pub fn validate(&self) -> Result<()> {
        self.state.bundle.validate()
    }

    /// Get document statistics after resolving the document sections.
    pub fn stats(&self) -> Result<PagesDocumentStats> {
        let total_objects = self.state.object_index.object_count();
        let section_count = self.state.document.section_count();

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
        assert!(doc.stats().unwrap().total_objects > 0);
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
        let sections = document.sections();
        assert_eq!(sections.len(), 1);
        assert_eq!(
            sections[0].text_storages()[0].text(),
            "Pages body — café 東京 🚀"
        );
        assert_eq!(sections[0].plain_text(), "Pages body — café 東京 🚀");
        assert_eq!(document.text().unwrap(), "Pages body — café 東京 🚀");
        assert_eq!(document.metadata().application.as_deref(), Some("Pages"));

        let snapshot = document.snapshot();
        assert!(std::ptr::eq(
            document.semantic_document(),
            snapshot.semantic_document()
        ));
        assert!(std::ptr::eq(
            document.sections().as_ptr(),
            document.semantic_document().sections().as_ptr()
        ));
        assert_eq!(
            snapshot.semantic_document().text_len(),
            sections[0].plain_text().len()
        );

        let structured = crate::structured::extract_sections(
            &document.state.bundle,
            &document.state.object_index,
        )
        .unwrap();
        assert_eq!(structured.len(), 1);
        assert_eq!(
            structured[0].text_storages()[0].text(),
            "Pages body — café 東京 🚀"
        );
    }
}
