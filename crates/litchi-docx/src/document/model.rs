//! Semantic values and state for the main WordprocessingML document.

use crate::alt::Chunk;
use crate::paragraph::Paragraph;
use crate::parts::DocumentPart;
use crate::table::Table;
use litchi_opc::OpcPackage;

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

/// An ordered main-document block containing a paragraph or table.
#[derive(Debug, Clone)]
pub enum Element {
    /// A paragraph block.
    Paragraph(Box<Paragraph>),
    /// A table block.
    Table(Box<Table>),
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
}

impl<'a> Document<'a> {
    /// Create a new Document from a DocumentPart and OpcPackage reference.
    ///
    /// This is typically called internally by `Package::document()`.
    #[inline]
    pub(crate) fn new(part: DocumentPart<'a>, opc: &'a OpcPackage) -> Self {
        Self { part, opc }
    }
}
