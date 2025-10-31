use super::super::{OleError, OleFile};
/// Package implementation for legacy PowerPoint presentations (.ppt).
use super::presentation::Presentation;
use std::fs::File;
use std::io::{self, Read, Seek};
use std::path::Path;

/// Error types for PPT file parsing.
#[derive(Debug)]
pub enum PptError {
    /// IO error
    Io(io::Error),
    /// OLE file error
    Ole(OleError),
    /// Invalid PPT format
    InvalidFormat(String),
    /// Stream not found
    StreamNotFound(String),
    /// Corrupted file
    Corrupted(String),
}

impl From<io::Error> for PptError {
    fn from(err: io::Error) -> Self {
        PptError::Io(err)
    }
}

impl From<OleError> for PptError {
    fn from(err: OleError) -> Self {
        PptError::Ole(err)
    }
}

impl std::fmt::Display for PptError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            PptError::Io(e) => write!(f, "IO error: {}", e),
            PptError::Ole(e) => write!(f, "OLE error: {}", e),
            PptError::InvalidFormat(s) => write!(f, "Invalid format: {}", s),
            PptError::StreamNotFound(s) => write!(f, "Stream not found: {}", s),
            PptError::Corrupted(s) => write!(f, "Corrupted file: {}", s),
        }
    }
}

impl std::error::Error for PptError {}

/// Result type for PPT operations.
pub type Result<T> = std::result::Result<T, PptError>;

/// A PowerPoint (.ppt) package.
///
/// This is the main entry point for working with legacy PowerPoint presentations.
/// It wraps an OLE file and provides PowerPoint-specific functionality.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ppt::Package;
///
/// // Open an existing presentation
/// let pkg = Package::open("presentation.ppt")?;
///
/// // Get the main presentation
/// let pres = pkg.presentation()?;
///
/// // Extract text
/// let text = pres.text()?;
/// println!("{}", text);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Package<R: Read + Seek = File> {
    /// The underlying OLE file
    ole: OleFile<R>,
}

impl Package<File> {
    /// Open a .ppt package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .ppt file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ppt::Package;
    ///
    /// let pkg = Package::open("presentation.ppt")?;
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
    /// * `reader` - A reader containing the .ppt file data
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use std::fs::File;
    /// use litchi::ole::ppt::Package;
    ///
    /// let file = File::open("presentation.ppt")?;
    /// let pkg = Package::from_reader(file)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_reader(reader: R) -> Result<Self> {
        let ole = OleFile::open(reader)?;

        // Verify it's a PowerPoint document by checking for the PowerPoint Document stream
        if !ole.exists(&["PowerPoint Document"]) {
            return Err(PptError::InvalidFormat(
                "Not a valid PowerPoint document: PowerPoint Document stream not found".to_string(),
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
    /// use litchi::ole::{OleFile, ppt::Package};
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("presentation.ppt")?;
    /// let ole = OleFile::open(Cursor::new(bytes))?;
    /// let pkg = Package::from_ole_file(ole)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_ole_file(ole: OleFile<R>) -> Result<Self> {
        // Verify it's a PowerPoint document by checking for the PowerPoint Document stream
        if !ole.exists(&["PowerPoint Document"]) {
            return Err(PptError::InvalidFormat(
                "Not a valid PowerPoint document: PowerPoint Document stream not found".to_string(),
            ));
        }

        Ok(Self { ole })
    }

    /// Get the main presentation.
    ///
    /// Returns the `Presentation` object which provides access to the presentation's
    /// content, slides, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ppt::Package;
    ///
    /// let pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn presentation(&mut self) -> Result<Presentation> {
        Presentation::from_ole(&mut self.ole)
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
    fn test_open_package() {
        let result = Package::open("test.ppt");
        assert!(result.is_ok());
    }

    #[test]
    #[ignore] // Requires test file
    fn test_invalid_file() {
        // Create a non-PPT file
        std::fs::write("test_invalid.tmp", b"Not a PPT file").unwrap();
        let result = Package::open("test_invalid.tmp");
        assert!(result.is_err());
        std::fs::remove_file("test_invalid.tmp").ok();
    }
}
