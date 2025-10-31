use super::super::{OleError, OleFile};
/// Package implementation for legacy Word documents (.doc).
use super::document::Document;
use std::fs::File;
use std::io::{self, Read, Seek};
use std::path::Path;

/// Error types for DOC file parsing.
#[derive(Debug)]
pub enum DocError {
    /// IO error
    Io(io::Error),
    /// OLE file error
    Ole(OleError),
    /// Invalid DOC format
    InvalidFormat(String),
    /// Stream not found
    StreamNotFound(String),
    /// Corrupted file
    Corrupted(String),
}

impl From<io::Error> for DocError {
    fn from(err: io::Error) -> Self {
        DocError::Io(err)
    }
}

impl From<OleError> for DocError {
    fn from(err: OleError) -> Self {
        DocError::Ole(err)
    }
}

impl std::fmt::Display for DocError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            DocError::Io(e) => write!(f, "IO error: {}", e),
            DocError::Ole(e) => write!(f, "OLE error: {}", e),
            DocError::InvalidFormat(s) => write!(f, "Invalid format: {}", s),
            DocError::StreamNotFound(s) => write!(f, "Stream not found: {}", s),
            DocError::Corrupted(s) => write!(f, "Corrupted file: {}", s),
        }
    }
}

impl std::error::Error for DocError {}

/// Result type for DOC operations.
pub type Result<T> = std::result::Result<T, DocError>;

/// A Word (.doc) package.
///
/// This is the main entry point for working with legacy Word documents.
/// It wraps an OLE file and provides Word-specific functionality.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::doc::Package;
///
/// // Open an existing document
/// let pkg = Package::open("document.doc")?;
///
/// // Get the main document
/// let doc = pkg.document()?;
///
/// // Extract text
/// let text = doc.text()?;
/// println!("{}", text);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Package<R: Read + Seek = File> {
    /// The underlying OLE file
    ole: OleFile<R>,
}

impl Package<File> {
    /// Open a .doc package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .doc file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::doc::Package;
    ///
    /// let pkg = Package::open("document.doc")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = File::open(path)?;
        Package::from_reader(file)
    }
}

impl<R: Read + Seek> Package<R> {
    /// Create a Package from any reader that implements Read + Seek.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader containing the .doc file data
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use std::fs::File;
    /// use litchi::ole::doc::Package;
    ///
    /// let file = File::open("document.doc")?;
    /// let pkg = Package::from_reader(file)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_reader(reader: R) -> Result<Self> {
        let ole = OleFile::open(reader)?;

        // Verify it's a Word document by checking for the WordDocument stream
        if !ole.exists(&["WordDocument"]) {
            return Err(DocError::InvalidFormat(
                "Not a valid Word document: WordDocument stream not found".to_string(),
            ));
        }

        Ok(Self { ole })
    }

    /// Create a Package from an already-parsed OLE file.
    ///
    /// This is used for single-pass parsing where the OLE file has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `ole` - An already-parsed OLE file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ole::{OleFile, doc::Package};
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("document.doc")?;
    /// let ole = OleFile::open(Cursor::new(bytes))?;
    /// let pkg = Package::from_ole_file(ole)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_ole_file(ole: OleFile<R>) -> Result<Self> {
        // Verify it's a Word document by checking for the WordDocument stream
        if !ole.exists(&["WordDocument"]) {
            return Err(DocError::InvalidFormat(
                "Not a valid Word document: WordDocument stream not found".to_string(),
            ));
        }

        Ok(Self { ole })
    }

    /// Get the main document.
    ///
    /// Returns the `Document` object which provides access to the document's
    /// content, formatting, tables, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::doc::Package;
    ///
    /// let pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn document(&mut self) -> Result<Document> {
        Document::from_ole(&mut self.ole)
    }

    /// Get the underlying OLE file.
    ///
    /// This provides access to lower-level OLE operations and streams.
    #[inline]
    pub fn ole_file(&mut self) -> &mut OleFile<R> {
        &mut self.ole
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    #[ignore] // Requires test file
    fn test_open_package() {
        let result = Package::open("test.doc");
        assert!(result.is_ok());
    }

    #[test]
    #[ignore] // Requires test file
    fn test_invalid_file() {
        // Create a non-DOC file
        std::fs::write("test_invalid.tmp", b"Not a DOC file").unwrap();
        let result = Package::open("test_invalid.tmp");
        assert!(result.is_err());
        std::fs::remove_file("test_invalid.tmp").ok();
    }
}
