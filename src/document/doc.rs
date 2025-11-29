//! Word document implementation.

use super::types::DocumentImpl;
use super::{Paragraph, Table};
use crate::common::{Error, Result};

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

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
        // Read file into memory and use smart detection for single-pass parsing
        // This is faster than the old approach of detecting first then parsing again
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes(bytes)
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
    /// - **Single-pass parsing**: Format detection reuses the parsed structure (40-60% faster)
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        // Use smart detection to parse only once
        use crate::common::detection::{DetectedFormat, detect_format_smart};

        let detected = detect_format_smart(bytes).ok_or(Error::NotOfficeFile)?;

        match detected {
            #[cfg(feature = "ole")]
            DetectedFormat::Doc(ole_file) => {
                // OLE file already parsed - reuse it!
                let mut package =
                    ole::doc::Package::from_ole_file(ole_file).map_err(Error::from)?;
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
            #[cfg(feature = "rtf")]
            DetectedFormat::Rtf(bytes) => {
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
            #[cfg(feature = "ooxml")]
            DetectedFormat::Docx(opc_package) => {
                // OPC package already parsed - reuse it!
                let package = Box::new(
                    ooxml::docx::Package::from_opc_package(opc_package).map_err(Error::from)?,
                );

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
            #[cfg(feature = "iwa")]
            DetectedFormat::Pages(data) => {
                let doc = crate::iwa::pages::PagesDocument::from_bytes(&data).map_err(|e| {
                    Error::ParseError(format!("Failed to open Pages document from bytes: {}", e))
                })?;

                Ok(Self {
                    inner: DocumentImpl::Pages(doc),
                    #[cfg(feature = "ooxml")]
                    _package: None,
                })
            },
            #[cfg(feature = "odf")]
            DetectedFormat::Odt(data) => {
                let doc = crate::odf::Document::from_bytes(data).map_err(|e| {
                    Error::ParseError(format!("Failed to parse ODT document from bytes: {}", e))
                })?;

                Ok(Self {
                    inner: DocumentImpl::Odt(doc),
                    #[cfg(feature = "ooxml")]
                    _package: None,
                })
            },
            // Handle mismatched formats
            #[allow(unreachable_patterns)]
            _ => Err(Error::InvalidFormat(
                "Detected format is not a document format or feature not enabled".to_string(),
            )),
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
            #[cfg(feature = "odf")]
            DocumentImpl::Odt(doc) => doc
                .text()
                .map_err(|e| Error::ParseError(format!("Failed to extract text from ODT: {}", e))),
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
            #[cfg(feature = "odf")]
            DocumentImpl::Odt(doc) => doc
                .paragraph_count()
                .map_err(|e| Error::ParseError(format!("Failed to get paragraph count: {}", e))),
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
                let paras = doc.paragraphs_with_content();
                // Convert to static lifetime by cloning the text
                let paras: Vec<_> = paras
                    .into_iter()
                    .map(|p| {
                        crate::rtf::ParagraphContent::new(
                            p.properties,
                            p.runs
                                .into_iter()
                                .map(|r| {
                                    crate::rtf::Run::new(
                                        std::borrow::Cow::Owned(r.text.into_owned()),
                                        r.formatting,
                                    )
                                })
                                .collect(),
                        )
                    })
                    .collect();
                Ok(paras.into_iter().map(Paragraph::Rtf).collect())
            },
            #[cfg(feature = "odf")]
            DocumentImpl::Odt(doc) => {
                let paras = doc
                    .paragraphs()
                    .map_err(|e| Error::ParseError(format!("Failed to get paragraphs: {}", e)))?;
                Ok(paras.into_iter().map(Paragraph::Odt).collect())
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
                Ok(tables
                    .into_iter()
                    .map(|t| Table::Docx(Box::new(t)))
                    .collect())
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
            #[cfg(feature = "odf")]
            DocumentImpl::Odt(doc) => {
                let tables = doc
                    .tables()
                    .map_err(|e| Error::ParseError(format!("Failed to get tables: {}", e)))?;
                Ok(tables.into_iter().map(Table::Odt).collect())
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
                        section.paragraphs.iter().map(|text| {
                            DocumentElement::Paragraph(Box::new(Paragraph::Pages(text.clone())))
                        })
                    })
                    .collect();
                Ok(elements)
            },
            #[cfg(feature = "rtf")]
            DocumentImpl::Rtf(doc) => {
                use super::DocumentElement;

                // Get elements from RTF document (paragraphs followed by tables)
                let rtf_elements = doc.elements();
                let mut elements = Vec::new();

                // Convert to owned elements with static lifetime
                for element in rtf_elements {
                    match element {
                        crate::rtf::DocumentElement::Paragraph(para) => {
                            let owned_para = crate::rtf::ParagraphContent::new(
                                para.properties,
                                para.runs
                                    .into_iter()
                                    .map(|r| {
                                        crate::rtf::Run::new(
                                            std::borrow::Cow::Owned(r.text.into_owned()),
                                            r.formatting,
                                        )
                                    })
                                    .collect(),
                            );
                            elements.push(DocumentElement::Paragraph(Box::new(Paragraph::Rtf(
                                owned_para,
                            ))));
                        },
                        crate::rtf::DocumentElement::Table(table) => {
                            let mut owned_table = crate::rtf::Table::new();
                            for row in table.rows() {
                                let mut owned_row = crate::rtf::Row::new();
                                for cell in row.cells() {
                                    let owned_cell = crate::rtf::Cell::new(
                                        std::borrow::Cow::Owned(cell.text().to_string()),
                                    );
                                    owned_row.add_cell(owned_cell);
                                }
                                owned_table.add_row(owned_row);
                            }
                            elements
                                .push(DocumentElement::Table(Box::new(Table::Rtf(owned_table))));
                        },
                    }
                }

                Ok(elements)
            },
            #[cfg(feature = "odf")]
            DocumentImpl::Odt(doc) => {
                use super::DocumentElement;
                use crate::odf::elements::parser::DocumentOrderElement;
                use crate::odf::elements::text::Paragraph as ElementParagraph;

                // Get ODF-specific elements and convert to unified API types
                let odf_elements = doc
                    .elements()
                    .map_err(|e| Error::ParseError(format!("Failed to get elements: {}", e)))?;

                let mut elements = Vec::new();
                for element in odf_elements {
                    match element {
                        DocumentOrderElement::Paragraph(para) => {
                            elements
                                .push(DocumentElement::Paragraph(Box::new(Paragraph::Odt(para))));
                        },
                        DocumentOrderElement::Heading(heading) => {
                            // Convert heading to paragraph for unified API
                            if let Ok(text) = heading.text() {
                                let mut para = ElementParagraph::new();
                                para.set_text(&text);
                                if let Some(style) = heading.style_name() {
                                    para.set_style_name(style);
                                }
                                elements.push(DocumentElement::Paragraph(Box::new(
                                    Paragraph::Odt(para),
                                )));
                            }
                        },
                        DocumentOrderElement::Table(table) => {
                            elements.push(DocumentElement::Table(Box::new(Table::Odt(table))));
                        },
                        DocumentOrderElement::List(_list) => {
                            // Lists are typically expanded to paragraphs in text extraction
                            // Skip in the unified document element API for now
                        },
                    }
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
            #[cfg(feature = "odf")]
            DocumentImpl::Odt(doc) => doc
                .metadata()
                .map_err(|e| Error::ParseError(format!("Failed to get metadata: {}", e))),
        }
    }
}
