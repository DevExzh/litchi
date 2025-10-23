//! Unified workbook implementation for Apple Numbers.

use super::types::Result;
use super::workbook_types::{
    WorkbookFormat, WorkbookImpl, detect_workbook_format_from_signature, refine_workbook_format,
};
use crate::common::{Error, Metadata};
use crate::sheet::WorkbookTrait;
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
    /// Cached metadata extracted during workbook initialization
    cached_metadata: Metadata,
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
        let file =
            File::open(path.as_ref()).map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
        let mut reader = BufReader::new(file);

        // Detect format
        let initial_format = detect_workbook_format_from_signature(&mut reader)
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;

        // Refine format for ZIP-based formats
        let format = refine_workbook_format(&mut reader, initial_format)
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;

        // Reset to beginning
        reader
            .seek(std::io::SeekFrom::Start(0))
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;

        // Open with appropriate implementation and extract metadata
        let (inner, metadata) = match format {
            #[cfg(feature = "iwa")]
            WorkbookFormat::Numbers => {
                let doc =
                    crate::iwa::numbers::NumbersDocument::open(path.as_ref()).map_err(|e| {
                        Box::new(Error::ParseError(format!("Failed to open Numbers: {}", e)))
                            as Box<dyn std::error::Error>
                    })?;

                // Extract metadata from Numbers bundle
                let metadata = Self::extract_numbers_metadata(&doc);
                (WorkbookImpl::Numbers(doc), metadata)
            },

            #[cfg(feature = "ole")]
            WorkbookFormat::Xls => {
                // Open file and extract metadata
                let file_for_metadata = File::open(path.as_ref())
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                let reader_for_metadata = std::io::BufReader::new(file_for_metadata);
                let mut ole_file = crate::ole::OleFile::open(reader_for_metadata)
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;

                let metadata = ole_file
                    .get_metadata()
                    .map(|m| m.into())
                    .unwrap_or_default();

                // Create XLS workbook with a fresh file reader
                let file = File::open(path.as_ref())
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                let reader = std::io::BufReader::new(file);
                let xls = crate::ole::xls::XlsWorkbook::new(reader)
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                (WorkbookImpl::XlsFile(xls), metadata)
            },

            #[cfg(feature = "ooxml")]
            WorkbookFormat::Xlsx => {
                let pkg = crate::ooxml::OpcPackage::open(path.as_ref())
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;

                // Extract metadata from OPC package
                let metadata = crate::ooxml::metadata::extract_metadata(&pkg).unwrap_or_default();

                let xlsx = crate::ooxml::xlsx::Workbook::new(pkg)?;
                (WorkbookImpl::Xlsx(xlsx), metadata)
            },

            #[cfg(feature = "ooxml")]
            WorkbookFormat::Xlsb => {
                // First, get metadata from the package
                let pkg_for_metadata = crate::ooxml::OpcPackage::open(path.as_ref())
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                let metadata =
                    crate::ooxml::metadata::extract_metadata(&pkg_for_metadata).unwrap_or_default();

                // Now create the XLSB workbook with a fresh reader
                let file = File::open(path.as_ref())
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                let reader = std::io::BufReader::new(file);
                let xlsb = crate::ooxml::xlsb::XlsbWorkbook::new(reader)
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                (WorkbookImpl::Xlsb(xlsb), metadata)
            },

            #[cfg(feature = "odf")]
            WorkbookFormat::Ods => {
                let ods = crate::odf::Spreadsheet::open(path.as_ref())
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                let metadata = ods.metadata().unwrap_or_default();
                (WorkbookImpl::Ods(std::cell::RefCell::new(ods)), metadata)
            },

            #[cfg(not(any(feature = "ole", feature = "ooxml", feature = "iwa", feature = "odf")))]
            _ => {
                return Err(Box::new(Error::ParseError(
                    "No workbook format support enabled".to_string(),
                )) as Box<dyn std::error::Error>);
            },

            // Fallback for unsupported combinations at compile-time
            #[allow(unreachable_patterns)]
            _ => {
                return Err(Box::new(Error::ParseError(
                    "Unsupported workbook format for current build features".to_string(),
                )) as Box<dyn std::error::Error>);
            },
        };

        Ok(Self {
            inner,
            cached_metadata: metadata,
        })
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

        // Open with appropriate implementation and extract metadata
        let (inner, metadata) = match format {
            #[cfg(feature = "iwa")]
            WorkbookFormat::Numbers => {
                let doc =
                    crate::iwa::numbers::NumbersDocument::from_bytes(&bytes).map_err(|e| {
                        Box::new(Error::ParseError(format!("Failed to parse Numbers: {}", e)))
                            as Box<dyn std::error::Error>
                    })?;

                // Extract metadata from Numbers bundle
                let metadata = Self::extract_numbers_metadata(&doc);
                (WorkbookImpl::Numbers(doc), metadata)
            },

            #[cfg(feature = "ole")]
            WorkbookFormat::Xls => {
                // Extract metadata from OLE file first
                let mut ole_file = crate::ole::OleFile::open(std::io::Cursor::new(&bytes))
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                let metadata = ole_file
                    .get_metadata()
                    .map(|m| m.into())
                    .unwrap_or_default();

                // Create XLS workbook with a fresh cursor
                let reader = std::io::Cursor::new(bytes);
                let xls = crate::ole::xls::XlsWorkbook::new(reader)
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                (WorkbookImpl::XlsMem(xls), metadata)
            },

            #[cfg(feature = "ooxml")]
            WorkbookFormat::Xlsx => {
                let pkg = crate::ooxml::OpcPackage::from_reader(std::io::Cursor::new(bytes))
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;

                // Extract metadata from OPC package
                let metadata = crate::ooxml::metadata::extract_metadata(&pkg).unwrap_or_default();

                let xlsx = crate::ooxml::xlsx::Workbook::new(pkg)?;
                (WorkbookImpl::Xlsx(xlsx), metadata)
            },

            #[cfg(feature = "ooxml")]
            WorkbookFormat::Xlsb => {
                // Extract metadata from package first
                let pkg_for_metadata =
                    crate::ooxml::OpcPackage::from_reader(std::io::Cursor::new(&bytes))
                        .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                let metadata =
                    crate::ooxml::metadata::extract_metadata(&pkg_for_metadata).unwrap_or_default();

                // Create XLSB workbook with a fresh cursor
                let xlsb = crate::ooxml::xlsb::XlsbWorkbook::new(std::io::Cursor::new(bytes))
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                (WorkbookImpl::Xlsb(xlsb), metadata)
            },

            #[cfg(feature = "odf")]
            WorkbookFormat::Ods => {
                let ods = crate::odf::Spreadsheet::from_bytes(bytes)
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                let metadata = ods.metadata().unwrap_or_default();
                (WorkbookImpl::Ods(std::cell::RefCell::new(ods)), metadata)
            },

            #[cfg(not(any(feature = "ole", feature = "ooxml", feature = "iwa", feature = "odf")))]
            _ => {
                return Err(Box::new(Error::ParseError(
                    "No workbook format support enabled".to_string(),
                )) as Box<dyn std::error::Error>);
            },

            // Fallback for unsupported combinations at compile-time
            #[allow(unreachable_patterns)]
            _ => {
                return Err(Box::new(Error::ParseError(
                    "Unsupported workbook format for current build features".to_string(),
                )) as Box<dyn std::error::Error>);
            },
        };

        Ok(Self {
            inner,
            cached_metadata: metadata,
        })
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
                let sheets = doc.sheets().map_err(|e| {
                    Box::new(Error::ParseError(format!("Failed to get sheets: {}", e)))
                        as Box<dyn std::error::Error>
                })?;
                Ok(sheets.iter().map(|s| s.name.clone()).collect())
            },

            #[cfg(feature = "ooxml")]
            WorkbookImpl::Xlsx(xlsx) => Ok(xlsx.worksheet_names().to_vec()),

            #[cfg(feature = "ooxml")]
            WorkbookImpl::Xlsb(xlsb) => Ok(xlsb.worksheet_names().to_vec()),

            #[cfg(feature = "ole")]
            WorkbookImpl::XlsFile(xls) => Ok(xls.worksheet_names().to_vec()),
            #[cfg(feature = "ole")]
            WorkbookImpl::XlsMem(xls) => Ok(xls.worksheet_names().to_vec()),

            #[cfg(feature = "odf")]
            WorkbookImpl::Ods(ods_ref) => {
                let mut ods = ods_ref.borrow_mut();
                let sheets = ods
                    .sheets()
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                Ok(sheets.iter().map(|s| s.name.clone()).collect())
            },

            #[cfg(any(feature = "ole", feature = "ooxml"))]
            WorkbookImpl::Other => Err(Box::new(Error::ParseError(
                "Unsupported workbook type in this build".to_string(),
            )) as Box<dyn std::error::Error>),
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
        match &self.inner {
            #[cfg(feature = "iwa")]
            WorkbookImpl::Numbers(doc) => {
                let sheets = doc
                    .sheets()
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                Ok(sheets.len())
            },
            #[cfg(feature = "ooxml")]
            WorkbookImpl::Xlsx(xlsx) => Ok(xlsx.worksheet_count()),
            #[cfg(feature = "ooxml")]
            WorkbookImpl::Xlsb(xlsb) => Ok(xlsb.worksheet_count()),
            #[cfg(feature = "ole")]
            WorkbookImpl::XlsFile(xls) => Ok(xls.worksheet_count()),
            #[cfg(feature = "ole")]
            WorkbookImpl::XlsMem(xls) => Ok(xls.worksheet_count()),
            #[cfg(feature = "odf")]
            WorkbookImpl::Ods(ods_ref) => {
                let mut ods = ods_ref.borrow_mut();
                let count = ods
                    .sheet_count()
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;
                Ok(count)
            },
            #[cfg(any(feature = "ole", feature = "ooxml"))]
            WorkbookImpl::Other => Err(Box::new(Error::ParseError(
                "Unsupported workbook type in this build".to_string(),
            )) as Box<dyn std::error::Error>),
        }
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
            WorkbookImpl::Numbers(doc) => doc.text().map_err(|e| {
                Box::new(Error::ParseError(format!(
                    "Failed to extract text from Numbers: {}",
                    e
                ))) as Box<dyn std::error::Error>
            }),

            #[cfg(feature = "ooxml")]
            WorkbookImpl::Xlsx(xlsx) => {
                // Iterate rows across worksheets
                let mut out = String::new();
                for i in 0..xlsx.worksheet_count() {
                    let ws = xlsx.worksheet_by_index(i)?;
                    let mut rows = ws.rows();
                    while let Some(row) = rows.next() {
                        let row = row?;
                        for (idx, cell) in row.iter().enumerate() {
                            if idx > 0 {
                                out.push('\t');
                            }
                            match cell {
                                crate::sheet::CellValue::Empty => {},
                                crate::sheet::CellValue::Bool(b) => {
                                    out.push_str(if *b { "TRUE" } else { "FALSE" })
                                },
                                crate::sheet::CellValue::Int(n) => out.push_str(&n.to_string()),
                                crate::sheet::CellValue::Float(f) => out.push_str(&f.to_string()),
                                crate::sheet::CellValue::String(s) => out.push_str(s),
                                crate::sheet::CellValue::DateTime(dt) => {
                                    out.push_str(&dt.to_string())
                                },
                                crate::sheet::CellValue::Error(e) => out.push_str(e),
                            }
                        }
                        out.push('\n');
                    }
                }
                Ok(out)
            },

            #[cfg(feature = "ooxml")]
            WorkbookImpl::Xlsb(xlsb) => {
                let mut out = String::new();
                for i in 0..xlsb.worksheet_count() {
                    let ws = xlsb.worksheet_by_index(i)?;
                    let mut rows = ws.rows();
                    while let Some(row) = rows.next() {
                        let row = row?;
                        for (idx, cell) in row.iter().enumerate() {
                            if idx > 0 {
                                out.push('\t');
                            }
                            match cell {
                                crate::sheet::CellValue::Empty => {},
                                crate::sheet::CellValue::Bool(b) => {
                                    out.push_str(if *b { "TRUE" } else { "FALSE" })
                                },
                                crate::sheet::CellValue::Int(n) => out.push_str(&n.to_string()),
                                crate::sheet::CellValue::Float(f) => out.push_str(&f.to_string()),
                                crate::sheet::CellValue::String(s) => out.push_str(s),
                                crate::sheet::CellValue::DateTime(dt) => {
                                    out.push_str(&dt.to_string())
                                },
                                crate::sheet::CellValue::Error(e) => out.push_str(e),
                            }
                        }
                        out.push('\n');
                    }
                }
                Ok(out)
            },

            #[cfg(feature = "ole")]
            WorkbookImpl::XlsFile(xls) => {
                let mut out = String::new();
                for i in 0..xls.worksheet_count() {
                    let ws = xls.worksheet_by_index(i)?;
                    let mut rows = ws.rows();
                    while let Some(row) = rows.next() {
                        let row = row?;
                        for (idx, cell) in row.iter().enumerate() {
                            if idx > 0 {
                                out.push('\t');
                            }
                            match cell {
                                crate::sheet::CellValue::Empty => {},
                                crate::sheet::CellValue::Bool(b) => {
                                    out.push_str(if *b { "TRUE" } else { "FALSE" })
                                },
                                crate::sheet::CellValue::Int(n) => out.push_str(&n.to_string()),
                                crate::sheet::CellValue::Float(f) => out.push_str(&f.to_string()),
                                crate::sheet::CellValue::String(s) => out.push_str(s),
                                crate::sheet::CellValue::DateTime(dt) => {
                                    out.push_str(&dt.to_string())
                                },
                                crate::sheet::CellValue::Error(e) => out.push_str(e),
                            }
                        }
                        out.push('\n');
                    }
                }
                Ok(out)
            },
            #[cfg(feature = "ole")]
            WorkbookImpl::XlsMem(xls) => {
                let mut out = String::new();
                for i in 0..xls.worksheet_count() {
                    let ws = xls.worksheet_by_index(i)?;
                    let mut rows = ws.rows();
                    while let Some(row) = rows.next() {
                        let row = row?;
                        for (idx, cell) in row.iter().enumerate() {
                            if idx > 0 {
                                out.push('\t');
                            }
                            match cell {
                                crate::sheet::CellValue::Empty => {},
                                crate::sheet::CellValue::Bool(b) => {
                                    out.push_str(if *b { "TRUE" } else { "FALSE" })
                                },
                                crate::sheet::CellValue::Int(n) => out.push_str(&n.to_string()),
                                crate::sheet::CellValue::Float(f) => out.push_str(&f.to_string()),
                                crate::sheet::CellValue::String(s) => out.push_str(s),
                                crate::sheet::CellValue::DateTime(dt) => {
                                    out.push_str(&dt.to_string())
                                },
                                crate::sheet::CellValue::Error(e) => out.push_str(e),
                            }
                        }
                        out.push('\n');
                    }
                }
                Ok(out)
            },

            #[cfg(feature = "odf")]
            WorkbookImpl::Ods(ods_ref) => {
                let mut ods = ods_ref.borrow_mut();
                ods.text()
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)
            },

            #[cfg(any(feature = "ole", feature = "ooxml"))]
            WorkbookImpl::Other => Err(Box::new(Error::ParseError(
                "Unsupported workbook type in this build".to_string(),
            )) as Box<dyn std::error::Error>),
        }
    }

    /// Get metadata from the workbook.
    ///
    /// Returns the cached metadata that was extracted during workbook initialization.
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
        Ok(self.cached_metadata.clone())
    }

    /// Extract metadata from a Numbers document.
    ///
    /// This extracts metadata from the Numbers bundle, similar to how
    /// Keynote metadata is extracted.
    #[cfg(feature = "iwa")]
    fn extract_numbers_metadata(doc: &crate::iwa::numbers::NumbersDocument) -> Metadata {
        let bundle_metadata = doc.bundle().metadata();
        let mut metadata = Metadata::default();

        // Extract title from properties
        if let Some(title) = bundle_metadata.get_property_string("Title") {
            metadata.title = Some(title);
        } else if let Some(title) = bundle_metadata.get_property_string("kDocumentTitleKey") {
            metadata.title = Some(title);
        }

        // Extract author
        if let Some(author) = bundle_metadata.get_property_string("Author") {
            metadata.author = Some(author);
        } else if let Some(author) = bundle_metadata.get_property_string("kDocumentAuthorKey") {
            metadata.author = Some(author);
        } else if let Some(author) = bundle_metadata.get_property_string("kSFWPAuthorPropertyKey") {
            metadata.author = Some(author);
        }

        // Extract keywords
        if let Some(keywords) = bundle_metadata.get_property_string("Keywords") {
            metadata.keywords = Some(keywords);
        }

        // Extract comments/description
        if let Some(comments) = bundle_metadata.get_property_string("Comments") {
            metadata.description = Some(comments);
        }

        // Extract application name
        if let Some(app) = bundle_metadata.detected_application.as_ref() {
            metadata.application = Some(app.clone());
        } else {
            metadata.application = Some("Numbers".to_string());
        }

        // Extract revision from Properties.plist
        if let Some(revision) = bundle_metadata.get_property_string("revision") {
            metadata.revision = Some(revision);
        }

        // Extract build version as additional version info
        if let Some(version) = bundle_metadata.latest_build_version()
            && metadata.revision.is_none()
        {
            metadata.revision = Some(version.to_string());
        }

        // Extract file format version
        if let Some(format_version) = bundle_metadata.get_property_string("fileFormatVersion") {
            metadata.content_status = Some(format!("Numbers Format Version {}", format_version));
        }

        metadata
    }
}
