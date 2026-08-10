//! Read-only, lazy DOCX access over an immutable positional source.
//!
//! [`Package::from_read_at`] validates the OPC package, its main-document
//! relationship, and the main-document content type without materializing the
//! main document. [`Package::document`] performs that first payload read and
//! returns a pinned semantic view which owns the loaded bytes.

use crate::error::Result;
use crate::package::validate_document_main_content_type;
use crate::paragraph::Paragraph;
use crate::parts::document_part::{document_paragraphs, visible_document_xml};
use litchi_core::ReadAt;
use litchi_opc::SourceBackedPackage;
use smallvec::SmallVec;
use std::sync::Arc;

/// A read-only DOCX package that leaves ordinary part bodies cold at open.
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
        Ok(self.paragraphs()?.len())
    }

    /// Return visible paragraphs sharing the pinned main-document allocation.
    pub fn paragraphs(&self) -> Result<SmallVec<[Paragraph; 32]>> {
        document_paragraphs(Arc::clone(&self.xml))
    }
}
