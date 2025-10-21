//! Unified workbook implementation for Apple Numbers.

use super::types::Result;
use super::workbook_types::{WorkbookImpl, WorkbookFormat, detect_workbook_format_from_signature, refine_workbook_format};
use crate::common::{Error, Metadata};
use std::fs::File;
use std::io::{BufReader, Cursor, Seek};
use std::path::Path;

/// A unified workbook interface for Apple Numbers spreadsheets.
///
/// This struct provides a high-level API for working with Apple Numbers files,
/// following the same pattern as the unified `Document` and `Presentation` APIs.
///
/// # Supported Formats
///
/// - `.numbers` - Apple Numbers (iWork Archive)
///
/// **Note**: For Excel formats (.xls, .xlsx, .xlsb), use the format-specific
/// APIs directly from `crate::ole::xls` or `crate::ooxml::xlsx`.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::sheet::Workbook;
///
/// // Open a Numbers spreadsheet
/// let workbook = Workbook::open("spreadsheet.numbers")?;
///
/// // Get worksheet names
/// let names = workbook.worksheet_names()?;
/// println!("Worksheets: {:?}", names);
///
/// // Extract all text
/// let text = workbook.text()?;
/// println!("{}", text);
///
/// // Get metadata
/// let metadata = workbook.metadata()?;
/// if let Some(title) = metadata.title {
///     println!("Title: {}", title);
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Workbook {
    inner: WorkbookImpl,
}

impl Workbook {
    /// Open a workbook from a file path.
    ///
    /// The format is automatically detected based on the file signature.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::sheet::Workbook;
    ///
    /// let workbook = Workbook::open("data.numbers")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = File::open(path.as_ref())
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
        let mut reader = BufReader::new(file);

        // Detect format
        let initial_format = detect_workbook_format_from_signature(&mut reader)
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
        
        // Refine format for ZIP-based formats
        let format = refine_workbook_format(&mut reader, initial_format)
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;

        // Reset to beginning
        reader.seek(std::io::SeekFrom::Start(0))
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;

        // Open with appropriate implementation
        let inner = match format {
            #[cfg(feature = "iwa")]
            WorkbookFormat::Numbers => {
                let doc = crate::iwa::numbers::NumbersDocument::open(path.as_ref())
                    .map_err(|e| Box::new(Error::ParseError(format!("Failed to open Numbers: {}", e))) as Box<dyn std::error::Error>)?;
                WorkbookImpl::Numbers(doc)
            }
            
            #[cfg(any(feature = "ole", feature = "ooxml"))]
            _ => {
                return Err(Box::new(Error::ParseError(
                    "This unified Workbook API currently only supports Apple Numbers. \
                     For Excel formats (.xls, .xlsx, .xlsb), use the format-specific APIs: \
                     crate::ole::xls::XlsWorkbook or crate::ooxml::xlsx::Workbook".to_string()
                )) as Box<dyn std::error::Error>);
            }
            
            #[cfg(not(any(feature = "ole", feature = "ooxml", feature = "iwa")))]
            _ => {
                return Err(Box::new(Error::ParseError("No workbook format support enabled".to_string())) as Box<dyn std::error::Error>);
            }
        };

        Ok(Self { inner })
    }

    /// Create a workbook from bytes.
    ///
    /// This is useful when you have the file data in memory.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::sheet::Workbook;
    /// use std::fs;
    ///
    /// let bytes = fs::read("data.numbers")?;
    /// let workbook = Workbook::from_bytes(bytes)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let mut cursor = Cursor::new(bytes.clone());

        // Detect format
        let initial_format = detect_workbook_format_from_signature(&mut cursor)
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
        
        // Refine format for ZIP-based formats
        let format = refine_workbook_format(&mut cursor, initial_format)
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;

        // Open with appropriate implementation
        let inner = match format {
            #[cfg(feature = "iwa")]
            WorkbookFormat::Numbers => {
                let doc = crate::iwa::numbers::NumbersDocument::from_bytes(&bytes)
                    .map_err(|e| Box::new(Error::ParseError(format!("Failed to parse Numbers: {}", e))) as Box<dyn std::error::Error>)?;
                WorkbookImpl::Numbers(doc)
            }
            
            #[cfg(any(feature = "ole", feature = "ooxml"))]
            _ => {
                return Err(Box::new(Error::ParseError(
                    "This unified Workbook API currently only supports Apple Numbers. \
                     For Excel formats (.xls, .xlsx, .xlsb), use the format-specific APIs".to_string()
                )) as Box<dyn std::error::Error>);
            }
            
            #[cfg(not(any(feature = "ole", feature = "ooxml", feature = "iwa")))]
            _ => {
                return Err(Box::new(Error::ParseError("No workbook format support enabled".to_string())) as Box<dyn std::error::Error>);
            }
        };

        Ok(Self { inner })
    }

    /// Get all worksheet names.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::sheet::Workbook;
    ///
    /// let workbook = Workbook::open("data.numbers")?;
    /// let names = workbook.worksheet_names()?;
    /// for name in names {
    ///     println!("Sheet: {}", name);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        match &self.inner {
            #[cfg(feature = "iwa")]
            WorkbookImpl::Numbers(doc) => {
                let sheets = doc.sheets()
                    .map_err(|e| Box::new(Error::ParseError(format!("Failed to get sheets: {}", e))) as Box<dyn std::error::Error>)?;
                Ok(sheets.iter().map(|s| s.name.clone()).collect())
            }
            
            #[cfg(any(feature = "ole", feature = "ooxml"))]
            WorkbookImpl::Other => {
                Err(Box::new(Error::ParseError("Not a Numbers workbook".to_string())) as Box<dyn std::error::Error>)
            }
        }
    }

    /// Get the number of worksheets.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::sheet::Workbook;
    ///
    /// let workbook = Workbook::open("data.numbers")?;
    /// println!("Number of sheets: {}", workbook.worksheet_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn worksheet_count(&self) -> Result<usize> {
        Ok(self.worksheet_names()?.len())
    }

    /// Extract all text from all worksheets.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::sheet::Workbook;
    ///
    /// let workbook = Workbook::open("data.numbers")?;
    /// let text = workbook.text()?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        match &self.inner {
            #[cfg(feature = "iwa")]
            WorkbookImpl::Numbers(doc) => {
                doc.text()
                    .map_err(|e| Box::new(Error::ParseError(format!("Failed to extract text from Numbers: {}", e))) as Box<dyn std::error::Error>)
            }
            
            #[cfg(any(feature = "ole", feature = "ooxml"))]
            WorkbookImpl::Other => {
                Err(Box::new(Error::ParseError("Not a Numbers workbook".to_string())) as Box<dyn std::error::Error>)
            }
        }
    }

    /// Get metadata from the workbook.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::sheet::Workbook;
    ///
    /// let workbook = Workbook::open("data.numbers")?;
    /// let metadata = workbook.metadata()?;
    /// if let Some(title) = metadata.title {
    ///     println!("Title: {}", title);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn metadata(&self) -> Result<Metadata> {
        match &self.inner {
            #[cfg(feature = "iwa")]
            WorkbookImpl::Numbers(_doc) => {
                // For now, return empty metadata for Numbers
                // TODO: Extract metadata from bundle when API is available
                Ok(Metadata::default())
            }
            
            #[cfg(any(feature = "ole", feature = "ooxml"))]
            WorkbookImpl::Other => {
                Ok(Metadata::default())
            }
        }
    }
}

