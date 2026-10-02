//! Semantic values and state for the main `WordprocessingML` document.

use crate::alt::Chunk;
use crate::paragraph::Paragraph;
use crate::parts::DocumentPart;
use crate::table::Table;
use litchi_opc::OpcPackage;
use std::sync::Arc;

/// A Word document.
///
/// This is the main API for reading and manipulating Word document content.
/// It provides access to paragraphs, tables, sections, styles, and other
/// document elements.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// // Extract all text
/// let text = doc.text()?;
/// println!("Document text: {}", text);
///
/// // Get paragraph count
/// let count = doc.paragraph_count()?;
/// println!("Number of paragraphs: {}", count);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Document<'a> {
    /// The underlying document part
    pub(super) part: DocumentPart<'a>,
    /// Reference to the OPC package (needed for accessing related parts like styles)
    pub(super) opc: &'a OpcPackage,
    /// Everything repaired while the document was read, aggregated from the
    /// codecs that record. The semantic text sink merges each
    /// `Document::write_text_to` call's repairs here; see `crate::leniency`.
    /// Interior mutability keeps the lazy read facades on `&self` (and the
    /// document `Send + Sync`) while still accumulating their repairs.
    pub(super) tolerance: std::sync::Mutex<crate::leniency::ToleranceReport>,
}

/// A picture watermark discovered in a document header, with its media part
/// resolved through the header part's relationships.
///
/// The payload is an inert borrowed byte view into the package; it is never
/// decoded, executed, or displayed.
#[derive(Debug)]
pub struct ImageWatermarkPart<'a> {
    /// Part name of the header carrying the watermark shape.
    pub source_header_name: String,
    /// Relationship ID of the `v:imagedata` reference in the header.
    pub relationship_id: String,
    /// Part name of the media part (e.g. `/word/media/watermarkImage1.png`).
    pub part_name: String,
    /// Declared OPC content type of the media part.
    pub content_type: &'a str,
    /// Original payload bytes held by the package.
    pub bytes: &'a [u8],
}

/// An active body child that this semantic facade does not model.
///
/// The exact selected XML element is retained without interpretation. It is
/// inert: callers can inspect the bytes but cannot execute, resolve, or mutate
/// the payload through this facade.
#[derive(Debug, Clone)]
pub struct OpaqueBlock {
    source: Arc<Vec<u8>>,
    start: u32,
    length: u32,
}

impl OpaqueBlock {
    pub(crate) const fn from_arc_range(source: Arc<Vec<u8>>, start: u32, length: u32) -> Self {
        Self {
            source,
            start,
            length,
        }
    }

    /// The retained active XML element with every namespace declaration it
    /// inherits from the document re-declared on its own root element, so that
    /// it parses standalone.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained range or the document around it
    /// cannot be read, or when a bound on the namespace scan is reached.
    pub fn self_contained_xml(&self) -> crate::error::Result<Vec<u8>> {
        crate::namespace::self_contained_element_xml(&self.source, self.start, self.length)
    }

    /// Borrow the retained active XML element exactly as selected by MCE.
    ///
    /// The bytes are an XML fragment whose namespace declarations are those in
    /// scope where it appears in the document, which since change 0653 are not
    /// repeated on the fragment's own root element. Use
    /// [`self_contained_xml`](Self::self_contained_xml) to parse it on its own.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        let Ok(start) = usize::try_from(self.start) else {
            return &[];
        };
        let Ok(length) = usize::try_from(self.length) else {
            return &[];
        };
        let Some(end) = start.checked_add(length) else {
            return &[];
        };
        self.source.get(start..end).unwrap_or_default()
    }
}

/// An ordered main-document block containing a paragraph or table.
#[derive(Debug, Clone)]
pub enum Element {
    /// A paragraph block.
    Paragraph(Box<Paragraph>),
    /// A table block.
    Table(Box<Table>),
    /// An active but unmodeled body child retained as exact XML.
    Unknown(Box<OpaqueBlock>),
}

/// An ordered main-document block, including inert alternative-format parts.
#[derive(Debug, Clone)]
pub enum Block {
    /// A paragraph block.
    Paragraph(Box<Paragraph>),
    /// A table block.
    Table(Box<Table>),
    /// An opaque alternative-format anchor.
    Alt(Box<Chunk>),
    /// An active but unmodeled body child retained as exact XML.
    Unknown(Box<OpaqueBlock>),
}

impl<'a> Document<'a> {
    /// Create a new Document from a `DocumentPart` and `OpcPackage` reference.
    ///
    /// This is typically called internally by `Package::document()`, which
    /// aggregates the repairs recorded while the document was opened into
    /// `tolerance`.
    #[inline]
    pub(crate) fn new(
        part: DocumentPart<'a>,
        opc: &'a OpcPackage,
        tolerance: crate::leniency::ToleranceReport,
    ) -> Self {
        Self {
            part,
            opc,
            tolerance: std::sync::Mutex::new(tolerance),
        }
    }

    /// Merge one read's repairs into the document-level report.
    pub(super) fn merge_tolerance(&self, repaired: &crate::leniency::ToleranceReport) {
        if repaired.is_clean() {
            return;
        }
        self.tolerance
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .append(repaired);
    }

    /// A snapshot of everything repaired while this document was read, in
    /// discovery order.
    ///
    /// The report is the exhaustive record of the non-structural substitutions
    /// described by [`crate::leniency`]: nothing is repaired silently. Reads
    /// merge their repairs here as their codecs run — the semantic text sink
    /// ([`Document::write_text_to`](crate::Document::write_text_to)) records
    /// foreign `*:t` skips — so the report accumulates across calls in call
    /// order and is empty only while every read conformed.
    #[must_use]
    pub fn tolerance_report(&self) -> crate::leniency::ToleranceReport {
        self.tolerance
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .clone()
    }
}
