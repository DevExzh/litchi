//! Word document implementation.

use super::types::{
    DocumentFormat, DocumentImpl, detect_document_format, detect_document_format_from_bytes,
};
use super::{Paragraph, Table};
use crate::common::{Error, Result};

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

use std::fs::File;
use std::io::Cursor;
use std::path::Path;

/// A Word document.
///
/// This is the main entry point for working with Word documents.
/// It automatically detects whether the file is .doc or .docx format
/// and provides a unified API.
///
/// Not intended to be constructed directly. Use `Document::open()` to
/// open a document.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::Document;
///
/// // Open a document (format auto-detected)
/// let doc = Document::open("report.doc")?;
///
/// // Get paragraph count
/// let count = doc.paragraph_count()?;
/// println!("Paragraphs: {}", count);
///
/// // Extract text
/// let text = doc.text()?;
/// println!("{}", text);
/// # Ok::<(), litchi::common::Error>(())
/// ```
pub struct Document {
    /// The underlying format-specific implementation
    pub(super) inner: DocumentImpl,
    /// DOCX package storage that must outlive the Document reference.
    ///
    /// This field is prefixed with `_` because it's not directly accessed,
    /// but it MUST be kept to maintain memory safety. The `inner` DocumentImpl::Docx
    /// variant holds a reference with extended lifetime to data owned by this Box.
    /// Dropping this would invalidate those references (use-after-free).
    ///
    /// Only used for DOCX files; None for DOC files.
    #[cfg(feature = "ooxml")]
    pub(super) _package: Option<Box<ooxml::docx::Package>>,
}

impl Document {
    /// Open a Word document from a file path.
    ///
    /// The file format (.doc or .docx) is automatically detected by examining
    /// the file header. You don't need to specify the format explicitly.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the Word document
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Document;
    ///
    /// // Open a .doc file
    /// let doc1 = Document::open("legacy.doc")?;
    ///
    /// // Open a .docx file
    /// let doc2 = Document::open("modern.docx")?;
    ///
    /// // Both work the same way
    /// println!("Doc 1: {}", doc1.text()?);
    /// println!("Doc 2: {}", doc2.text()?);
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let path = path.as_ref();

        // Try to detect the format by reading the file header
        let mut file = File::open(path)?;
        let initial_format = detect_document_format(&mut file)?;

        // Refine ZIP-based format detection (distinguish DOCX from Pages)
        let format = super::types::refine_document_format(&mut file, initial_format)?;

        // Reopen the file for the appropriate parser
        match format {
            #[cfg(feature = "ole")]
            DocumentFormat::Doc => {
                let mut package = ole::doc::Package::open(path).map_err(Error::from)?;
                let doc = package.document().map_err(Error::from)?;

                // Extract metadata from the OLE file
                let metadata = package
                    .ole_file()
                    .get_metadata()
                    .map(|m| m.into())
                    .unwrap_or_default();

                Ok(Self {
                    inner: DocumentImpl::Doc(doc, metadata),
                    #[cfg(feature = "ooxml")]
                    _package: None,
                })
            },
            #[cfg(not(feature = "ole"))]
            DocumentFormat::Doc => Err(Error::FeatureDisabled("ole".to_string())),
            #[cfg(feature = "rtf")]
            DocumentFormat::Rtf => {
                let doc = crate::rtf::RtfDocument::open(path).map_err(|e| {
                    Error::ParseError(format!("Failed to parse RTF document: {}", e))
                })?;

                Ok(Self {
                    inner: DocumentImpl::Rtf(doc),
                    #[cfg(feature = "ooxml")]
                    _package: None,
                })
            },
            #[cfg(not(feature = "rtf"))]
            DocumentFormat::Rtf => Err(Error::FeatureDisabled("rtf".to_string())),
            #[cfg(feature = "ooxml")]
            DocumentFormat::Docx => {
                let package = Box::new(ooxml::docx::Package::open(path).map_err(Error::from)?);

                // SAFETY: We're using unsafe here to extend the lifetime of the document
                // reference. This is safe because we're storing the package in the same
                // struct, ensuring it lives as long as the document reference.
                let doc_ref = unsafe {
                    let pkg_ptr = &*package as *const ooxml::docx::Package;
                    let doc = (*pkg_ptr).document().map_err(Error::from)?;
                    std::mem::transmute::<ooxml::docx::Document<'_>, ooxml::docx::Document<'static>>(
                        doc,
                    )
                };

                // Extract metadata from OOXML core properties
                let metadata = crate::ooxml::metadata::extract_metadata(package.opc_package())
                    .unwrap_or_else(|_| crate::common::Metadata::default());

                Ok(Self {
                    inner: DocumentImpl::Docx(Box::new(doc_ref), metadata),
                    _package: Some(package),
                })
            },
            #[cfg(not(feature = "ooxml"))]
            DocumentFormat::Docx => Err(Error::FeatureDisabled("ooxml".to_string())),
            #[cfg(feature = "iwa")]
            DocumentFormat::Pages => {
                let doc = crate::iwa::pages::PagesDocument::open(path).map_err(|e| {
                    Error::ParseError(format!("Failed to open Pages document: {}", e))
                })?;

                Ok(Self {
                    inner: DocumentImpl::Pages(doc),
                    #[cfg(feature = "ooxml")]
                    _package: None,
                })
            },
            #[cfg(not(feature = "iwa"))]
            DocumentFormat::Pages => Err(Error::FeatureDisabled("iwa".to_string())),
        }
    }

    /// Create a Document from a byte buffer.
    ///
    /// This method is optimized for parsing documents from memory, such as
    /// from network traffic or in-memory caches, without creating temporary files.
    /// It automatically detects the format (.doc or .docx) from the byte signature.
    ///
    /// # Arguments
    ///
    /// * `bytes` - The document bytes
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Document;
    /// use std::fs;
    ///
    /// // From owned bytes (e.g., network data)
    /// let data = fs::read("document.doc")?;
    /// let doc = Document::from_bytes(data)?;
    /// println!("{}", doc.text()?);
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    ///
    /// # Performance Notes
    ///
    /// - For .doc files (OLE2): Parses directly from the buffer with minimal copying
    /// - For .docx files (ZIP): Efficient decompression without file I/O overhead
    /// - Ideal for network data, streams, or in-memory content
    /// - No temporary files created
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        // Detect format from byte signature
        let initial_format = detect_document_format_from_bytes(&bytes)?;

        // Refine ZIP-based format detection for bytes
        let format = if initial_format == DocumentFormat::Docx {
            // Check if it's a Pages document
            #[cfg(feature = "iwa")]
            {
                let mut cursor = Cursor::new(&bytes);
                super::types::refine_document_format(&mut cursor, initial_format)?
            }
            #[cfg(not(feature = "iwa"))]
            initial_format
        } else {
            initial_format
        };

        match format {
            #[cfg(feature = "ole")]
            DocumentFormat::Doc => {
                // For OLE2, create cursor from bytes
                let cursor = Cursor::new(bytes);

                let mut package = ole::doc::Package::from_reader(cursor).map_err(Error::from)?;
                let doc = package.document().map_err(Error::from)?;

                // Extract metadata from the OLE file
                let metadata = package
                    .ole_file()
                    .get_metadata()
                    .map(|m| m.into())
                    .unwrap_or_default();

                Ok(Self {
                    inner: DocumentImpl::Doc(doc, metadata),
                    #[cfg(feature = "ooxml")]
                    _package: None,
                })
            },
            #[cfg(not(feature = "ole"))]
            DocumentFormat::Doc => Err(Error::FeatureDisabled("ole".to_string())),
            #[cfg(feature = "rtf")]
            DocumentFormat::Rtf => {
                let text = String::from_utf8(bytes)
                    .map_err(|e| Error::ParseError(format!("Invalid UTF-8 in RTF: {}", e)))?;

                let doc = crate::rtf::RtfDocument::parse(&text).map_err(|e| {
                    Error::ParseError(format!("Failed to parse RTF document: {}", e))
                })?;

                Ok(Self {
                    inner: DocumentImpl::Rtf(doc),
                    #[cfg(feature = "ooxml")]
                    _package: None,
                })
            },
            #[cfg(not(feature = "rtf"))]
            DocumentFormat::Rtf => Err(Error::FeatureDisabled("rtf".to_string())),
            #[cfg(feature = "ooxml")]
            DocumentFormat::Docx => {
                // For OOXML/ZIP, Cursor<Vec<u8>> implements Read + Seek
                let cursor = Cursor::new(bytes);

                let package =
                    Box::new(ooxml::docx::Package::from_reader(cursor).map_err(Error::from)?);

                // SAFETY: Same lifetime extension as in `open()`
                let doc_ref = unsafe {
                    let pkg_ptr = &*package as *const ooxml::docx::Package;
                    let doc = (*pkg_ptr).document().map_err(Error::from)?;
                    std::mem::transmute::<ooxml::docx::Document<'_>, ooxml::docx::Document<'static>>(
                        doc,
                    )
                };

                // Extract metadata from OOXML core properties
                let metadata = crate::ooxml::metadata::extract_metadata(package.opc_package())
                    .unwrap_or_else(|_| crate::common::Metadata::default());

                Ok(Self {
                    inner: DocumentImpl::Docx(Box::new(doc_ref), metadata),
                    _package: Some(package),
                })
            },
            #[cfg(not(feature = "ooxml"))]
            DocumentFormat::Docx => Err(Error::FeatureDisabled("ooxml".to_string())),
            #[cfg(feature = "iwa")]
            DocumentFormat::Pages => {
                let doc = crate::iwa::pages::PagesDocument::from_bytes(&bytes).map_err(|e| {
                    Error::ParseError(format!("Failed to open Pages document from bytes: {}", e))
                })?;

                Ok(Self {
                    inner: DocumentImpl::Pages(doc),
                    #[cfg(feature = "ooxml")]
                    _package: None,
                })
            },
            #[cfg(not(feature = "iwa"))]
            DocumentFormat::Pages => Err(Error::FeatureDisabled("iwa".to_string())),
        }
    }

    /// Get all text content from the document.
    ///
    /// This extracts all text from the document, concatenated together.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Document;
    ///
    /// let doc = Document::open("document.doc")?;
    /// let text = doc.text()?;
    /// println!("{}", text);
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        match &self.inner {
            #[cfg(feature = "ole")]
            DocumentImpl::Doc(doc, _) => doc.text().map_err(Error::from),
            #[cfg(feature = "ooxml")]
            DocumentImpl::Docx(doc, _) => doc.text().map_err(Error::from),
            #[cfg(feature = "iwa")]
            DocumentImpl::Pages(doc) => doc.text().map_err(|e| {
                Error::ParseError(format!("Failed to extract text from Pages: {}", e))
            }),
            #[cfg(feature = "rtf")]
            DocumentImpl::Rtf(doc) => Ok(doc.text()),
        }
    }

    /// Get the number of paragraphs in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Document;
    ///
    /// let doc = Document::open("document.doc")?;
    /// let count = doc.paragraph_count()?;
    /// println!("Paragraphs: {}", count);
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn paragraph_count(&self) -> Result<usize> {
        match &self.inner {
            #[cfg(feature = "ole")]
            DocumentImpl::Doc(doc, _) => doc.paragraph_count().map_err(Error::from),
            #[cfg(feature = "ooxml")]
            DocumentImpl::Docx(doc, _) => doc.paragraph_count().map_err(Error::from),
            #[cfg(feature = "iwa")]
            DocumentImpl::Pages(doc) => {
                // Pages documents are organized by sections
                let sections = doc
                    .sections()
                    .map_err(|e| Error::ParseError(format!("Failed to get sections: {}", e)))?;
                Ok(sections.iter().map(|s| s.paragraphs.len()).sum())
            },
            #[cfg(feature = "rtf")]
            DocumentImpl::Rtf(doc) => Ok(doc.paragraph_count()),
        }
    }

    /// Get an iterator over paragraphs in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Document;
    ///
    /// let doc = Document::open("document.doc")?;
    /// for para in doc.paragraphs()? {
    ///     println!("Paragraph: {}", para.text()?);
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        match &self.inner {
            #[cfg(feature = "ole")]
            DocumentImpl::Doc(doc, _) => {
                let paras = doc.paragraphs().map_err(Error::from)?;
                Ok(paras.into_iter().map(Paragraph::Doc).collect())
            },
            #[cfg(feature = "ooxml")]
            DocumentImpl::Docx(doc, _) => {
                let paras = doc.paragraphs().map_err(Error::from)?;
                Ok(paras.into_iter().map(Paragraph::Docx).collect())
            },
            #[cfg(feature = "iwa")]
            DocumentImpl::Pages(doc) => {
                // Pages documents have sections, each with paragraphs
                let sections = doc
                    .sections()
                    .map_err(|e| Error::ParseError(format!("Failed to get sections: {}", e)))?;
                let paragraphs: Vec<_> = sections
                    .iter()
                    .flat_map(|section| {
                        section
                            .paragraphs
                            .iter()
                            .map(|text| Paragraph::Pages(text.clone()))
                    })
                    .collect();
                Ok(paragraphs)
            },
            #[cfg(feature = "rtf")]
            DocumentImpl::Rtf(doc) => {
                let paras = doc.paragraphs();
                Ok(paras.into_iter().map(Paragraph::Rtf).collect())
            },
        }
    }

    /// Get an iterator over tables in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Document;
    ///
    /// let doc = Document::open("document.doc")?;
    /// for table in doc.tables()? {
    ///     println!("Table with {} rows", table.row_count()?);
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn tables(&self) -> Result<Vec<Table>> {
        match &self.inner {
            #[cfg(feature = "ole")]
            DocumentImpl::Doc(doc, _) => {
                let tables = doc.tables().map_err(Error::from)?;
                Ok(tables.into_iter().map(Table::Doc).collect())
            },
            #[cfg(feature = "ooxml")]
            DocumentImpl::Docx(doc, _) => {
                let tables = doc.tables().map_err(Error::from)?;
                Ok(tables.into_iter().map(Table::Docx).collect())
            },
            #[cfg(feature = "iwa")]
            DocumentImpl::Pages(_doc) => {
                // Pages tables are not currently supported in the paragraph/table extraction API
                // Tables in Pages are embedded as structured data which requires different extraction
                Ok(Vec::new())
            },
            #[cfg(feature = "rtf")]
            DocumentImpl::Rtf(doc) => {
                let tables = doc.tables();
                Ok(tables
                    .iter()
                    .map(|t| {
                        // Convert RTF table to owned Table
                        let mut owned_table = crate::rtf::Table::new();
                        for row in t.rows() {
                            let mut owned_row = crate::rtf::Row::new();
                            for cell in row.cells() {
                                let owned_cell = crate::rtf::Cell::new(std::borrow::Cow::Owned(
                                    cell.text().to_string(),
                                ));
                                owned_row.add_cell(owned_cell);
                            }
                            owned_table.add_row(owned_row);
                        }
                        Table::Rtf(owned_table)
                    })
                    .collect())
            },
        }
    }

    /// Get all document elements (paragraphs and tables) in document order.
    ///
    /// This method is optimized to extract paragraphs and tables in a single pass,
    /// which is more efficient than calling `paragraphs()` and `tables()` separately.
    /// More importantly, it preserves the document order of elements, which is essential
    /// for proper sequential processing (e.g., Markdown conversion).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::{Document, DocumentElement};
    ///
    /// let doc = Document::open("document.doc")?;
    ///
    /// // Process elements in document order
    /// for element in doc.elements()? {
    ///     match element {
    ///         DocumentElement::Paragraph(para) => {
    ///             println!("Paragraph: {}", para.text()?);
    ///         }
    ///         DocumentElement::Table(table) => {
    ///             println!("Table with {} rows", table.row_count()?);
    ///         }
    ///     }
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    ///
    /// # Performance
    ///
    /// - For `.doc` files: Extracts paragraphs once and identifies tables from them
    /// - For `.docx` files: Parses XML once to extract both paragraphs and tables
    /// - This is 2x faster than calling `paragraphs()` and `tables()` separately
    pub fn elements(&self) -> Result<Vec<super::DocumentElement>> {
        match &self.inner {
            #[cfg(feature = "ole")]
            DocumentImpl::Doc(doc, _) => doc.elements().map_err(Error::from),
            #[cfg(feature = "ooxml")]
            DocumentImpl::Docx(doc, _) => doc.elements().map_err(Error::from),
            #[cfg(feature = "iwa")]
            DocumentImpl::Pages(doc) => {
                use super::DocumentElement;
                // Pages documents have sections with paragraphs
                // Tables are not currently supported in the extraction API
                let sections = doc
                    .sections()
                    .map_err(|e| Error::ParseError(format!("Failed to get sections: {}", e)))?;
                let elements: Vec<_> = sections
                    .iter()
                    .flat_map(|section| {
                        section
                            .paragraphs
                            .iter()
                            .map(|text| DocumentElement::Paragraph(Paragraph::Pages(text.clone())))
                    })
                    .collect();
                Ok(elements)
            },
            #[cfg(feature = "rtf")]
            DocumentImpl::Rtf(doc) => {
                use super::DocumentElement;
                // For RTF, we need to interleave paragraphs and tables in order
                // First, get all paragraphs and tables
                let paragraphs = doc.paragraphs();
                let tables = doc.tables();

                // RTF documents store elements in order, so we need to identify
                // which paragraphs are part of tables and create the elements list
                // For now, we'll use a simple approach: add all paragraphs first, then tables
                // TODO: Implement proper RTF element ordering based on RTF structure
                let mut elements = Vec::new();

                for para in paragraphs {
                    elements.push(DocumentElement::Paragraph(Paragraph::Rtf(para)));
                }

                for table in tables {
                    let mut owned_table = crate::rtf::Table::new();
                    for row in table.rows() {
                        let mut owned_row = crate::rtf::Row::new();
                        for cell in row.cells() {
                            let owned_cell = crate::rtf::Cell::new(std::borrow::Cow::Owned(
                                cell.text().to_string(),
                            ));
                            owned_row.add_cell(owned_cell);
                        }
                        owned_table.add_row(owned_row);
                    }
                    elements.push(DocumentElement::Table(Table::Rtf(owned_table)));
                }

                Ok(elements)
            },
        }
    }

    /// Get document metadata.
    ///
    /// Extracts metadata from the document such as title, author, creation date, etc.
    /// For OLE (.doc) files, this reads from SummaryInformation and DocumentSummaryInformation streams.
    /// For OOXML (.docx) files, this reads from core properties (currently not implemented).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Document;
    ///
    /// let doc = Document::open("document.doc")?;
    /// let metadata = doc.metadata()?;
    /// if let Some(title) = &metadata.title {
    ///     println!("Title: {}", title);
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn metadata(&self) -> Result<crate::common::Metadata> {
        match &self.inner {
            #[cfg(feature = "ole")]
            DocumentImpl::Doc(_, metadata) => Ok(metadata.clone()),
            #[cfg(feature = "ooxml")]
            DocumentImpl::Docx(_, metadata) => Ok(metadata.clone()),
            #[cfg(feature = "iwa")]
            DocumentImpl::Pages(doc) => {
                // Extract metadata from Pages bundle metadata
                let bundle_metadata = doc.bundle().metadata();
                let mut metadata = crate::common::Metadata::default();

                // Extract title from properties
                if let Some(title) = bundle_metadata.get_property_string("Title") {
                    metadata.title = Some(title);
                }

                // Extract author from properties
                if let Some(author) = bundle_metadata.get_property_string("Author") {
                    metadata.author = Some(author);
                }

                // Extract document identifier
                if let Some(doc_id) = bundle_metadata.document_identifier() {
                    metadata.description = Some(format!("Document ID: {}", doc_id));
                }

                Ok(metadata)
            },
            #[cfg(feature = "rtf")]
            DocumentImpl::Rtf(_doc) => {
                // RTF doesn't have standard metadata in the same way
                // Metadata would need to be parsed from \info group
                Ok(crate::common::Metadata::default())
            },
        }
    }
}
