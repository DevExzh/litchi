/// Unified Word document module.
///
/// This module provides a unified API for working with Word documents in both
/// legacy (.doc) and modern (.docx) formats. The format is automatically detected
/// and handled transparently.
///
/// # Architecture
///
/// The module provides a format-agnostic API following the python-docx design:
/// - `Document`: The main document API (auto-detects format)
/// - `Paragraph`: Paragraph with text runs
/// - `Run`: Text run with formatting
/// - `Table`: Table with rows and cells
///
/// # Example
///
/// ```rust,no_run
/// use litchi::Document;
///
/// // Open any Word document (.doc or .docx) - format auto-detected
/// let doc = Document::open("document.doc")?;
///
/// // Extract all text
/// let text = doc.text()?;
/// println!("Document text: {}", text);
///
/// // Access paragraphs
/// for para in doc.paragraphs()? {
///     println!("Paragraph: {}", para.text()?);
///
///     // Access runs in paragraph
///     for run in para.runs()? {
///         println!("  Run: {} (bold: {:?})", run.text()?, run.bold()?);
///     }
/// }
///
/// // Access tables
/// for table in doc.tables()? {
///     for row in table.rows()? {
///         for cell in row.cells()? {
///             println!("Cell: {}", cell.text()?);
///         }
///     }
/// }
/// # Ok::<(), litchi::common::Error>(())
/// ```

use crate::common::{Error, Result};
use crate::ole;
use crate::ooxml;
use std::fs::File;
use std::io::{Cursor, Read, Seek};
use std::path::Path;

/// A Word document that can be either .doc or .docx format.
///
/// This enum wraps the format-specific implementations and provides
/// a unified API. Users typically don't interact with this enum directly,
/// but instead use the methods on `Document`.
enum DocumentImpl {
    /// Legacy .doc format
    Doc(ole::doc::Document),
    /// Modern .docx format  
    Docx(Box<ooxml::docx::Document<'static>>),
}

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
    inner: DocumentImpl,
    /// DOCX package storage that must outlive the Document reference.
    ///
    /// This field is prefixed with `_` because it's not directly accessed,
    /// but it MUST be kept to maintain memory safety. The `inner` DocumentImpl::Docx
    /// variant holds a reference with extended lifetime to data owned by this Box.
    /// Dropping this would invalidate those references (use-after-free).
    ///
    /// Only used for DOCX files; None for DOC files.
    _package: Option<Box<ooxml::docx::Package>>,
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
        let format = detect_document_format(&mut file)?;
        
        // Reopen the file for the appropriate parser
        match format {
            DocumentFormat::Doc => {
                let mut package = ole::doc::Package::open(path)
                    .map_err(Error::from)?;
                let doc = package.document()
                    .map_err(Error::from)?;
                
                Ok(Self {
                    inner: DocumentImpl::Doc(doc),
                    _package: None,
                })
            }
            DocumentFormat::Docx => {
                let package = Box::new(ooxml::docx::Package::open(path)
                    .map_err(Error::from)?);
                
                // SAFETY: We're using unsafe here to extend the lifetime of the document
                // reference. This is safe because we're storing the package in the same
                // struct, ensuring it lives as long as the document reference.
                let doc_ref = unsafe {
                    let pkg_ptr = &*package as *const ooxml::docx::Package;
                    let doc = (*pkg_ptr).document()
                        .map_err(Error::from)?;
                    std::mem::transmute::<ooxml::docx::Document<'_>, ooxml::docx::Document<'static>>(doc)
                };
                
                Ok(Self {
                    inner: DocumentImpl::Docx(Box::new(doc_ref)),
                    _package: Some(package),
                })
            }
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
        let format = detect_document_format_from_bytes(&bytes)?;
        
        match format {
            DocumentFormat::Doc => {
                // For OLE2, create cursor from bytes
                let cursor = Cursor::new(bytes);
                
                let mut package = ole::doc::Package::from_reader(cursor)
                    .map_err(Error::from)?;
                let doc = package.document()
                    .map_err(Error::from)?;
                
                Ok(Self {
                    inner: DocumentImpl::Doc(doc),
                    _package: None,
                })
            }
            DocumentFormat::Docx => {
                // For OOXML/ZIP, Cursor<Vec<u8>> implements Read + Seek
                let cursor = Cursor::new(bytes);
                
                let package = Box::new(ooxml::docx::Package::from_reader(cursor)
                    .map_err(Error::from)?);
                
                // SAFETY: Same lifetime extension as in `open()`
                let doc_ref = unsafe {
                    let pkg_ptr = &*package as *const ooxml::docx::Package;
                    let doc = (*pkg_ptr).document()
                        .map_err(Error::from)?;
                    std::mem::transmute::<ooxml::docx::Document<'_>, ooxml::docx::Document<'static>>(doc)
                };
                
                Ok(Self {
                    inner: DocumentImpl::Docx(Box::new(doc_ref)),
                    _package: Some(package),
                })
            }
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
            DocumentImpl::Doc(doc) => {
                doc.text().map_err(Error::from)
            }
            DocumentImpl::Docx(doc) => {
                doc.text().map_err(Error::from)
            }
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
            DocumentImpl::Doc(doc) => {
                doc.paragraph_count().map_err(Error::from)
            }
            DocumentImpl::Docx(doc) => {
                doc.paragraph_count().map_err(Error::from)
            }
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
            DocumentImpl::Doc(doc) => {
                let paras = doc.paragraphs()
                    .map_err(Error::from)?;
                Ok(paras.into_iter().map(Paragraph::Doc).collect())
            }
            DocumentImpl::Docx(doc) => {
                let paras = doc.paragraphs()
                    .map_err(Error::from)?;
                Ok(paras.into_iter().map(Paragraph::Docx).collect())
            }
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
            DocumentImpl::Doc(doc) => {
                let tables = doc.tables()
                    .map_err(Error::from)?;
                Ok(tables.into_iter().map(Table::Doc).collect())
            }
            DocumentImpl::Docx(doc) => {
                let tables = doc.tables()
                    .map_err(Error::from)?;
                Ok(tables.into_iter().map(Table::Docx).collect())
            }
        }
    }
}

/// A paragraph in a Word document.
pub enum Paragraph {
    Doc(ole::doc::Paragraph),
    Docx(ooxml::docx::Paragraph),
}

impl Paragraph {
    /// Get the text content of the paragraph.
    pub fn text(&self) -> Result<String> {
        match self {
            Paragraph::Doc(p) => p.text().map(|s| s.to_string()).map_err(Error::from),
            Paragraph::Docx(p) => p.text().map(|s| s.to_string()).map_err(Error::from),
        }
    }

    /// Get the runs in this paragraph.
    pub fn runs(&self) -> Result<Vec<Run>> {
        match self {
            Paragraph::Doc(p) => {
                let runs = p.runs().map_err(Error::from)?;
                Ok(runs.into_iter().map(Run::Doc).collect())
            }
            Paragraph::Docx(p) => {
                let runs = p.runs().map_err(Error::from)?;
                Ok(runs.into_iter().map(Run::Docx).collect())
            }
        }
    }
}

/// A text run in a paragraph.
pub enum Run {
    Doc(ole::doc::Run),
    Docx(ooxml::docx::Run),
}

impl Run {
    /// Get the text content of the run.
    pub fn text(&self) -> Result<String> {
        match self {
            Run::Doc(r) => r.text().map(|s| s.to_string()).map_err(Error::from),
            Run::Docx(r) => r.text().map(|s| s.to_string()).map_err(Error::from),
        }
    }

    /// Check if the run is bold.
    pub fn bold(&self) -> Result<Option<bool>> {
        match self {
            Run::Doc(r) => Ok(r.bold()),
            Run::Docx(r) => r.bold().map_err(Error::from),
        }
    }

    /// Check if the run is italic.
    pub fn italic(&self) -> Result<Option<bool>> {
        match self {
            Run::Doc(r) => Ok(r.italic()),
            Run::Docx(r) => r.italic().map_err(Error::from),
        }
    }
}

/// A table in a Word document.
pub enum Table {
    Doc(ole::doc::Table),
    Docx(ooxml::docx::Table),
}

impl Table {
    /// Get the number of rows in the table.
    pub fn row_count(&self) -> Result<usize> {
        match self {
            Table::Doc(t) => t.row_count().map_err(Error::from),
            Table::Docx(t) => t.row_count().map_err(Error::from),
        }
    }

    /// Get the rows in this table.
    pub fn rows(&self) -> Result<Vec<Row>> {
        match self {
            Table::Doc(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.into_iter().map(Row::Doc).collect())
            }
            Table::Docx(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.into_iter().map(Row::Docx).collect())
            }
        }
    }
}

/// A table row in a Word document.
pub enum Row {
    Doc(ole::doc::Row),
    Docx(ooxml::docx::Row),
}

impl Row {
    /// Get the cells in this row.
    pub fn cells(&self) -> Result<Vec<Cell>> {
        match self {
            Row::Doc(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.into_iter().map(Cell::Doc).collect())
            }
            Row::Docx(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.into_iter().map(Cell::Docx).collect())
            }
        }
    }
}

/// A table cell in a Word document.
pub enum Cell {
    Doc(ole::doc::Cell),
    Docx(ooxml::docx::Cell),
}

impl Cell {
    /// Get the text content of the cell.
    pub fn text(&self) -> Result<String> {
        match self {
            Cell::Doc(c) => c.text().map(|s| s.to_string()).map_err(Error::from),
            Cell::Docx(c) => c.text().map(|s| s.to_string()).map_err(Error::from),
        }
    }
}

/// Document format detection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DocumentFormat {
    /// Legacy .doc format (OLE2)
    Doc,
    /// Modern .docx format (OOXML/ZIP)
    Docx,
}

/// Detect the document format by reading the file header.
///
/// This function reads the first few bytes of the file to determine if it's
/// an OLE2 file (.doc) or a ZIP file (.docx).
fn detect_document_format<R: Read + Seek>(reader: &mut R) -> Result<DocumentFormat> {
    use std::io::SeekFrom;

    // Read the first 8 bytes
    let mut header = [0u8; 8];
    reader.read_exact(&mut header)?;
    
    // Reset to the beginning
    reader.seek(SeekFrom::Start(0))?;

    detect_document_format_from_signature(&header)
}

/// Detect the document format from a byte buffer.
///
/// This is optimized for in-memory detection without seeking.
#[inline]
fn detect_document_format_from_bytes(bytes: &[u8]) -> Result<DocumentFormat> {
    if bytes.len() < 4 {
        return Err(Error::InvalidFormat("File too small to determine format".to_string()));
    }
    
    detect_document_format_from_signature(&bytes[0..8.min(bytes.len())])
}

/// Detect format from the signature bytes.
#[inline]
fn detect_document_format_from_signature(header: &[u8]) -> Result<DocumentFormat> {
    // Check for OLE2 signature (D0 CF 11 E0 A1 B1 1A E1)
    if header.len() >= 4 && header[0..4] == [0xD0, 0xCF, 0x11, 0xE0] {
        return Ok(DocumentFormat::Doc);
    }

    // Check for ZIP signature (PK\x03\x04)
    if header.len() >= 4 && header[0..4] == [0x50, 0x4B, 0x03, 0x04] {
        return Ok(DocumentFormat::Docx);
    }

    Err(Error::NotOfficeFile)
}

