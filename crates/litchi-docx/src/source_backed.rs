//! Lazy DOCX access and guarded main-document publication over an immutable
//! positional source.
//!
//! [`Package::from_read_at`] validates the OPC package, its main-document
//! relationship, and the main-document content type without materializing the
//! main document. [`Package::document`] performs that first payload read and
//! returns a pinned semantic view which owns the loaded bytes. Main-document
//! transactions retain the raw XML and may be published to a sequential sink
//! while raw-copying every unselected ZIP member.

use crate::document::{Commit, Edit, Snapshot, TransactionResult};
use crate::error::{Error, Result};
use crate::package::validate_document_main_content_type;
use crate::paragraph::Paragraph;
use crate::parts::document_part::{
    document_paragraph, document_paragraph_count, document_paragraphs, visible_document_xml,
};
use litchi_core::ReadAt;
use litchi_opc::SourceBackedPackage;
use smallvec::SmallVec;
use std::io::Write;
use std::sync::Arc;

/// A DOCX package that leaves ordinary part bodies cold at open.
pub struct Package {
    package: SourceBackedPackage,
}

impl Package {
    /// Open a DOCX source using the standard bounded OPC read policy.
    ///
    /// This validates the main-document relationship and content type but
    /// does not decompress or materialize the main-document payload.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_source_backed(SourceBackedPackage::from_read_at(source)?)
    }

    /// Open a DOCX source with an explicit bounded OPC read policy.
    ///
    /// This validates the main-document relationship and content type but
    /// does not decompress or materialize the main-document payload.
    pub fn from_read_at_with_limits(
        source: Arc<dyn ReadAt>,
        limits: litchi_opc::ReadLimits,
    ) -> Result<Self> {
        Self::from_source_backed(SourceBackedPackage::from_read_at_with_limits(
            source, limits,
        )?)
    }

    fn from_source_backed(package: SourceBackedPackage) -> Result<Self> {
        let main = package.main_document_part()?;
        validate_document_main_content_type(main.content_type())?;
        Ok(Self { package })
    }

    /// Load and pin the main document for read-only semantic queries.
    ///
    /// The first call reads only the main-document part. The returned document
    /// owns its normalized XML bytes, so repeated text and paragraph queries do
    /// not revisit the positional source.
    pub fn document(&self) -> Result<Document> {
        let main = self.package.main_document_part()?;
        validate_document_main_content_type(main.content_type())?;
        let xml = visible_document_xml(main.data()?.into_arc())?;
        Ok(Document { xml })
    }

    /// Return content-free payload-cache activity for this lazy package.
    ///
    /// This does not read any part payload or expose member identities.
    #[must_use]
    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.package.cache_diagnostics()
    }

    /// Load the exact raw main-document bytes as a semantic transaction
    /// snapshot.
    ///
    /// Source-backed edits currently refuse documents whose markup-
    /// compatibility preprocessing selects or rewrites branches. Keeping the
    /// raw bytes prevents semantic selectors from being applied to a different
    /// XML representation than the one eventually published.
    ///
    /// # Errors
    ///
    /// Returns a typed package, document, resource-limit, or unsafe-edit error.
    pub fn document_snapshot(&self) -> TransactionResult<Snapshot> {
        let (_, snapshot) = self.main_document_snapshot("document_snapshot")?;
        Ok(snapshot)
    }

    /// Start an isolated semantic edit over the exact raw main document.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::document_snapshot`].
    pub fn edit_document(&self) -> TransactionResult<Edit> {
        Ok(self.document_snapshot()?.edit())
    }

    /// Publish one exact-source-checked main-document commit to a sequential
    /// stream while preserving every other physical ZIP member.
    ///
    /// Only operations confined to the main-document payload are accepted.
    /// Cross-package paragraph transfers are refused because their dependency
    /// graph requires package-level publication. A no-op commit reproduces the
    /// complete source artifact byte for byte. A changed signed package is
    /// refused by the underlying OPC publisher.
    ///
    /// All semantic, source-version, topology, signature, and replacement-XML
    /// checks happen before output. A sink failure after output begins is
    /// reported through the underlying typed incomplete-output error.
    ///
    /// # Errors
    ///
    /// Returns a typed transaction, package, unsafe-edit, signature, source,
    /// XML-publication, or sink error.
    pub fn publish_document_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &Commit,
    ) -> TransactionResult<Snapshot> {
        let (main, current) = self.main_document_snapshot("publish_document_commit_to_stream")?;
        let target = commit.patch().apply(&current)?;
        if commit
            .patch()
            .operations()
            .iter()
            .any(|operation| !operation.supports_source_backed_main_document_overlay())
        {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "publish_document_commit_to_stream",
                reason: "paragraph transfers require package-level dependency publication",
            }
            .into());
        }
        self.package
            .write_part_overlay_to_stream(writer, &main, target.xml_bytes().to_vec())
            .map_err(Error::from)?;
        Ok(target)
    }

    fn main_document_snapshot(
        &self,
        operation: &'static str,
    ) -> TransactionResult<(litchi_opc::PackURI, Snapshot)> {
        let main = self.package.main_document_part().map_err(Error::from)?;
        validate_document_main_content_type(main.content_type())?;
        let partname = main.partname().clone();
        let raw = main.data().map_err(Error::from)?.into_arc();
        let visible = visible_document_xml(Arc::clone(&raw))?;
        if !Arc::ptr_eq(&raw, &visible) {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "source-backed document transactions do not support markup-compatibility branch selection",
            }
            .into());
        }
        Ok((partname, Snapshot::from_shared_xml(raw)?))
    }
}

/// A pinned read-only view of a DOCX main document.
///
/// This view owns the main-document bytes loaded by [`Package::document`].
/// It intentionally exposes semantic text and paragraph queries only; use
/// the established [`crate::Package`] APIs for mutable package access.
#[derive(Clone)]
pub struct Document {
    xml: Arc<Vec<u8>>,
}

impl Document {
    /// Extract all visible paragraph text from the pinned document.
    pub fn extract_text(&self) -> Result<String> {
        crate::paragraph::extract_word_text(self.xml.as_slice())
    }

    /// Count visible paragraphs in the pinned document.
    pub fn paragraph_count(&self) -> Result<usize> {
        document_paragraph_count(self.xml.as_slice())
    }

    /// Return visible paragraphs sharing the pinned main-document allocation.
    pub fn paragraphs(&self) -> Result<SmallVec<[Paragraph; 32]>> {
        document_paragraphs(Arc::clone(&self.xml))
    }

    /// Return one visible paragraph without allocating all paragraph views.
    pub fn paragraph(&self, index: usize) -> Result<Option<Paragraph>> {
        document_paragraph(Arc::clone(&self.xml), index)
    }
}
