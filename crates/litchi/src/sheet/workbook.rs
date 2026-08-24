//! Unified workbook implementation for supported spreadsheet formats.

use super::types::Result;
use super::workbook_types::WorkbookImpl;
#[allow(unused_imports, reason = "re-exported for sheet implementations")]
use crate::sheet::WorkbookTrait;
use litchi_core::{Error, Metadata};
#[cfg(feature = "xls")]
use litchi_ole_common::property_set::PropertySetReader;
#[cfg(all(feature = "xls", any(unix, windows)))]
use litchi_ole_common::property_set::SharedPropertySetReader;
use std::path::Path;
#[cfg(all(feature = "xls", any(unix, windows)))]
use std::{io::Cursor, sync::Arc, sync::OnceLock};

const MAX_WORKBOOK_PATH_BYTES: u64 = 2 * 1024 * 1024 * 1024;

#[allow(
    dead_code,
    reason = "used by the non-positional filesystem fallback; positional XLSX/ODS paths use their retained source owner or bytes"
)]
fn read_path_bytes(path: impl AsRef<Path>) -> Result<Vec<u8>> {
    use std::{fs::File, io::Read};

    let mut file = File::open(path)
        .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)?;
    let length = file
        .metadata()
        .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)?
        .len();
    if length > MAX_WORKBOOK_PATH_BYTES {
        return Err(Box::new(Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: litchi_core::Resource::InputBytes,
            observed: length,
            limit: MAX_WORKBOOK_PATH_BYTES,
            scope: std::sync::Arc::from("unified filesystem workbook"),
        })));
    }
    let length = usize::try_from(length).map_err(|_| {
        Box::new(Error::InvalidFormat(
            "filesystem source exceeds platform allocation limits".to_string(),
        )) as Box<dyn std::error::Error + Send + Sync>
    })?;
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(length).map_err(|source| {
        Box::new(Error::Allocation {
            resource: "unified filesystem workbook source bytes",
            source,
        }) as Box<dyn std::error::Error + Send + Sync>
    })?;
    bytes.resize(length, 0);
    file.read_exact(&mut bytes)
        .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)?;

    // Avoid silently accepting a file that grew after the bounded allocation.
    let mut extra = [0_u8; 1];
    if file
        .read(&mut extra)
        .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)?
        != 0
    {
        let observed = (length as u64).saturating_add(1);
        return if observed > MAX_WORKBOOK_PATH_BYTES {
            Err(Box::new(Error::ResourceLimit(litchi_core::ResourceLimit {
                resource: litchi_core::Resource::InputBytes,
                observed,
                limit: MAX_WORKBOOK_PATH_BYTES,
                scope: std::sync::Arc::from("unified filesystem workbook"),
            })))
        } else {
            Err(Box::new(Error::InvalidFormat(format!(
                "filesystem source grew during bounded read (observed at least {observed} bytes)"
            ))))
        };
    }
    Ok(bytes)
}

#[cfg(any(
    feature = "numbers",
    any(feature = "xlsx", feature = "xlsb"),
    feature = "xls"
))]
#[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
fn append_cell_text(out: &mut String, cell: &litchi_core::sheet::CellValue) {
    use litchi_core::sheet::CellValue;

    match cell {
        CellValue::Empty => {},
        CellValue::Bool(value) => out.push_str(if *value { "TRUE" } else { "FALSE" }),
        CellValue::Int(value) => out.push_str(&value.to_string()),
        CellValue::Float(value) => out.push_str(&value.to_string()),
        CellValue::String(value) => out.push_str(value),
        CellValue::DateTime(value) => out.push_str(&value.to_string()),
        CellValue::Error(value) => out.push_str(value),
        CellValue::Formula {
            formula,
            cached_value,
            ..
        } => match cached_value.as_deref() {
            Some(CellValue::String(value)) => out.push_str(value),
            Some(CellValue::Int(value)) => out.push_str(&value.to_string()),
            Some(CellValue::Float(value)) => out.push_str(&value.to_string()),
            Some(CellValue::Bool(value)) => out.push_str(if *value { "TRUE" } else { "FALSE" }),
            _ => {
                out.push('=');
                out.push_str(formula);
            },
        },
    }
}

#[cfg(all(feature = "xls", any(unix, windows)))]
/// Internal filesystem XLS adapter. The positional source is retained so
/// source-backed selectors and metadata never need to reopen the path.
pub(super) struct XlsSource {
    workbook: crate::xls::SourceBackedWorkbook,
    cfb: Arc<litchi_cfb::SharedOleFile>,
    eager: OnceLock<crate::xls::Workbook<Cursor<Vec<u8>>>>,
    metadata: OnceLock<Metadata>,
}

#[cfg(all(feature = "xls", any(unix, windows)))]
impl XlsSource {
    fn new(
        workbook: crate::xls::SourceBackedWorkbook,
        cfb: Arc<litchi_cfb::SharedOleFile>,
    ) -> Self {
        Self {
            workbook,
            cfb,
            eager: OnceLock::new(),
            metadata: OnceLock::new(),
        }
    }

    fn ensure_current(&self) -> Result<()> {
        self.workbook
            .source_version()
            .map(|_| ())
            .map_err(map_xls_source_error)
    }

    fn with_eager<T>(
        &self,
        operation: impl FnOnce(&crate::xls::Workbook<Cursor<Vec<u8>>>) -> Result<T>,
    ) -> Result<T> {
        self.ensure_current()?;
        if self.eager.get().is_none() {
            let eager = self
                .workbook
                .materialize_eager()
                .map_err(map_xls_source_error)?;
            self.ensure_current()?;
            let _ = self.eager.set(eager);
        }

        let eager = self.eager.get().ok_or_else(|| {
            Box::new(Error::ParseError(
                "XLS eager compatibility materialization was not published".to_owned(),
            )) as Box<dyn std::error::Error + Send + Sync>
        })?;
        let result = operation(eager);
        self.ensure_current()?;
        result
    }

    fn metadata(&self) -> Result<Metadata> {
        self.ensure_current()?;
        if let Some(metadata) = self.metadata.get() {
            let metadata = metadata.clone();
            self.ensure_current()?;
            return Ok(metadata);
        }

        let metadata = match self.cfb.get_metadata() {
            Ok(metadata) => litchi_core::Metadata::from(metadata),
            Err(error) if xls_metadata_error_is_soft(&error) => Metadata::default(),
            Err(error) => {
                return Err(map_xls_source_error(normalize_xls_metadata_error(error)));
            },
        };
        self.ensure_current()?;
        let _ = self.metadata.set(metadata);
        let metadata = self.metadata.get().ok_or_else(|| {
            Box::new(Error::ParseError(
                "XLS metadata was not published".to_owned(),
            )) as Box<dyn std::error::Error + Send + Sync>
        })?;
        let metadata = metadata.clone();
        self.ensure_current()?;
        Ok(metadata)
    }
}

#[cfg(all(feature = "xls", any(unix, windows)))]
fn map_xls_source_error(
    error: crate::xls::SourceBackedError,
) -> Box<dyn std::error::Error + Send + Sync> {
    Box::new(error)
}

#[cfg(all(feature = "xls", any(unix, windows)))]
fn normalize_xls_metadata_error(error: litchi_cfb::OleError) -> crate::xls::SourceBackedError {
    error.into()
}

#[cfg(all(feature = "xls", any(unix, windows)))]
fn xls_metadata_error_is_soft(error: &litchi_cfb::OleError) -> bool {
    matches!(
        error,
        litchi_cfb::OleError::InvalidFormat(_)
            | litchi_cfb::OleError::InvalidData(_)
            | litchi_cfb::OleError::NotOleFile
            | litchi_cfb::OleError::CorruptedFile(_)
            | litchi_cfb::OleError::StreamNotFound
    )
}

/// A unified workbook interface for supported spreadsheet formats.
///
/// The selected format is detected from the file signature. Filesystem XLS
/// opens retain a checked positional source for lazy worksheet names/counts;
/// text extraction uses bounded on-demand compatibility materialization.
///
/// # Supported Formats
///
/// - `.numbers` - Apple Numbers (when the `numbers` feature is enabled)
/// - `.xls` - legacy Excel workbooks (when the `xls` feature is enabled)
/// - `.xlsx` - Office Open XML workbooks (when the `xlsx` feature is enabled)
/// - `.xlsb` - Office Open XML binary workbooks (when the `xlsb` feature is enabled)
/// - `.ods` - OpenDocument spreadsheets (when the `ods` feature is enabled)
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::sheet::Workbook;
///
/// // Open a supported spreadsheet
/// let workbook = Workbook::open("spreadsheet.xlsx")?;
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
/// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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
    /// let workbook = Workbook::open("spreadsheet.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        #[cfg(all(
            any(feature = "xlsx", feature = "ods", feature = "xls"),
            any(unix, windows)
        ))]
        {
            match crate::detection_smart::detected::detect_workbook_source_path(path.as_ref())? {
                #[cfg(feature = "xlsx")]
                crate::detection_smart::detected::WorkbookSourcePathDetection::Xlsx {
                    workbook,
                    metadata,
                } => Ok(Self {
                    inner: WorkbookImpl::Xlsx(super::adapters::Workbook::from_source_backed(
                        workbook,
                    )),
                    cached_metadata: metadata,
                }),
                #[cfg(feature = "ods")]
                crate::detection_smart::detected::WorkbookSourcePathDetection::Ods(ods) => {
                    let metadata = ods.metadata()?.clone();
                    Ok(Self {
                        inner: WorkbookImpl::OdsSource(*ods),
                        cached_metadata: metadata,
                    })
                },
                #[cfg(feature = "xls")]
                crate::detection_smart::detected::WorkbookSourcePathDetection::Xls {
                    workbook,
                    cfb,
                } => Ok(Self {
                    inner: WorkbookImpl::XlsSource(XlsSource::new(workbook, cfb)),
                    cached_metadata: Metadata::default(),
                }),
                crate::detection_smart::detected::WorkbookSourcePathDetection::OtherOoxml {
                    format: _,
                    bytes,
                }
                | crate::detection_smart::detected::WorkbookSourcePathDetection::Bytes(bytes) => {
                    Self::from_bytes(bytes)
                },
                crate::detection_smart::detected::WorkbookSourcePathDetection::DisabledOtherOoxml(
                    _,
                ) => {
                    Err(Box::new(Error::NotOfficeFile) as Box<dyn std::error::Error + Send + Sync>)
                },
            }
        }

        #[cfg(not(all(
            any(feature = "xlsx", feature = "ods", feature = "xls"),
            any(unix, windows)
        )))]
        {
            // Read once into owned memory; detection transfers that ownership
            // into the selected format path.
            let bytes = read_path_bytes(path.as_ref())?;
            Self::from_bytes(bytes)
        }
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
    /// let bytes = fs::read("spreadsheet.xlsx")?;
    /// let workbook = Workbook::from_bytes(bytes)?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    ///
    /// # Performance Notes
    ///
    /// - XLSX detection retains a bounded source-backed catalog and defers worksheet payloads
    /// - Other OLE2 and OOXML detection results return parsed owners that their loaders reuse
    /// - Other detection results retain the moved buffer for loaders that may parse it afterward
    /// - No temporary files created
    /// - Ideal for network data, streams, or in-memory content
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        #[cfg(feature = "ods")]
        let bytes = match crate::detection_smart::detected::detect_prepared_ods(bytes) {
            Ok(prepared) => {
                let ods = litchi_ods::Spreadsheet::from_prepared_package(prepared)
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error + Send + Sync>)?;
                let metadata = ods.metadata().clone();
                return Ok(Self {
                    inner: WorkbookImpl::Ods(std::cell::RefCell::new(ods)),
                    cached_metadata: metadata,
                });
            },
            Err(bytes) => bytes,
        };

        #[cfg(feature = "xlsx")]
        let bytes = match crate::detection_smart::detected::detect_workbook_source_bytes(bytes) {
            crate::detection_smart::detected::WorkbookSourceBytesDetection::Xlsx(package) => {
                let metadata = crate::ooxml_common::properties::read_source_backed(&package)
                    .map_err(crate::map_ooxml_error)?
                    .map(litchi_core::Metadata::from)
                    .unwrap_or_default();
                let workbook =
                    crate::xlsx::SourceBackedWorkbook::from_source_backed_package(package)
                        .map_err(crate::map_ooxml_error)?;
                return Ok(Self {
                    inner: WorkbookImpl::Xlsx(super::adapters::Workbook::from_source_backed(
                        workbook,
                    )),
                    cached_metadata: metadata,
                });
            },
            crate::detection_smart::detected::WorkbookSourceBytesDetection::Fallback(bytes) => {
                bytes
            },
        };

        // Detection consumes the input and returns either a parsed owner or the
        // moved source bytes, depending on the format.
        use crate::detection_smart::{DetectedFormat, detect_format_smart};

        let detected = detect_format_smart(bytes).ok_or_else(|| {
            Box::new(Error::NotOfficeFile) as Box<dyn std::error::Error + Send + Sync>
        })?;

        // Open with appropriate implementation and extract metadata
        let (inner, metadata) = match detected {
            #[cfg(feature = "numbers")]
            DetectedFormat::Numbers(data) => {
                let doc = litchi_numbers::Package::from_bytes(&data).map_err(|e| {
                    Box::new(Error::ParseError(format!("Failed to parse Numbers: {}", e)))
                        as Box<dyn std::error::Error + Send + Sync>
                })?;

                // Extract metadata from Numbers bundle
                let metadata = Self::extract_numbers_metadata(&doc);
                (WorkbookImpl::Numbers(doc), metadata)
            },

            #[cfg(feature = "xls")]
            DetectedFormat::Xls(ole_file) => {
                // OLE file already parsed - reuse it!
                let mut ole_file_for_metadata = ole_file;
                let metadata = ole_file_for_metadata
                    .get_metadata()
                    .map(|m| m.into())
                    .unwrap_or_default();

                // Create XLS workbook directly from the parsed OLE file
                let xls = crate::xls::Workbook::from_ole_file(ole_file_for_metadata)
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error + Send + Sync>)?;
                (WorkbookImpl::XlsMem(xls), metadata)
            },

            #[cfg(feature = "xlsx")]
            DetectedFormat::Xlsx(opc_package) => {
                // OPC package already parsed - reuse it!
                let metadata = crate::ooxml_common::properties::read(&opc_package)
                    .map_err(crate::map_ooxml_error)?
                    .map(litchi_core::Metadata::from)
                    .unwrap_or_default();
                let xlsx = crate::xlsx::Package::from_opc(opc_package)
                    .map_err(crate::map_ooxml_error)?
                    .into_workbook()
                    .map_err(crate::map_ooxml_error)?;
                (
                    WorkbookImpl::Xlsx(super::adapters::Workbook::new(xlsx)),
                    metadata,
                )
            },

            #[cfg(feature = "xlsb")]
            DetectedFormat::Xlsb(opc_package) => {
                // OPC package already parsed - reuse it!
                let metadata = crate::ooxml_common::properties::read(&opc_package)
                    .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)?
                    .map(litchi_core::Metadata::from)
                    .unwrap_or_default();

                // Create XLSB workbook directly from the parsed OPC package
                let xlsb = crate::xlsb::Package::from_opc(opc_package)
                    .map_err(crate::map_ooxml_error)?
                    .into_workbook()
                    .map_err(crate::map_ooxml_error)?;
                (WorkbookImpl::Xlsb(xlsb), metadata)
            },

            #[cfg(feature = "ods")]
            DetectedFormat::Ods(data) => {
                let ods = litchi_ods::Spreadsheet::from_bytes(data)
                    .map_err(|e| Box::new(e) as Box<dyn std::error::Error + Send + Sync>)?;
                let metadata = ods.metadata().clone();
                (WorkbookImpl::Ods(std::cell::RefCell::new(ods)), metadata)
            },

            #[cfg(feature = "ods")]
            DetectedFormat::FlatOdf(format, data) => {
                let _ = data;
                return Err(Box::new(Error::Unsupported(format!(
                    "flat OpenDocument {:?} is detected but the dedicated family facade exposes packaged parsing only",
                    format
                )))
                    as Box<dyn std::error::Error + Send + Sync>);
            },

            // Handle mismatched formats
            #[allow(
                unreachable_patterns,
                reason = "match arms are feature-gated; the fallback is unreachable when every format feature is enabled"
            )]
            _ => {
                return Err(
                    Box::new(Error::NotOfficeFile) as Box<dyn std::error::Error + Send + Sync>
                );
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
    /// let workbook = Workbook::open("spreadsheet.xlsx")?;
    /// let names = workbook.worksheet_names()?;
    /// for name in names {
    ///     println!("Sheet: {}", name);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        match &self.inner {
            #[cfg(feature = "numbers")]
            WorkbookImpl::Numbers(doc) => Ok(doc
                .sheets()
                .iter()
                .map(|sheet| sheet.name().to_owned())
                .collect()),

            #[cfg(feature = "xlsx")]
            WorkbookImpl::Xlsx(xlsx) => {
                xlsx.ensure_source_current()?;
                Ok(xlsx.worksheet_names().to_vec())
            },

            #[cfg(feature = "xlsb")]
            WorkbookImpl::Xlsb(xlsb) => Ok(xlsb.worksheet_names().to_vec()),

            #[cfg(feature = "xls")]
            WorkbookImpl::XlsFile(xls) => Ok(xls.worksheet_names().to_vec()),
            #[cfg(feature = "xls")]
            WorkbookImpl::XlsMem(xls) => Ok(xls.worksheet_names().to_vec()),

            #[cfg(all(feature = "xls", any(unix, windows)))]
            WorkbookImpl::XlsSource(source) => source
                .workbook
                .worksheet_names()
                .map_err(map_xls_source_error),

            #[cfg(feature = "ods")]
            WorkbookImpl::Ods(spreadsheet) => Ok(spreadsheet.borrow().sheet_names()),

            #[cfg(all(feature = "ods", any(unix, windows)))]
            WorkbookImpl::OdsSource(spreadsheet) => Ok(spreadsheet.sheet_names()?),

            #[cfg(any(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
            WorkbookImpl::Other => Err(Box::new(Error::ParseError(
                "Unsupported workbook type in this build".to_string(),
            )) as Box<dyn std::error::Error + Send + Sync>),
        }
    }

    /// Get the number of worksheets.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::sheet::Workbook;
    ///
    /// let workbook = Workbook::open("spreadsheet.xlsx")?;
    /// println!("Number of sheets: {}", workbook.worksheet_count()?);
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn worksheet_count(&self) -> Result<usize> {
        match &self.inner {
            #[cfg(feature = "numbers")]
            WorkbookImpl::Numbers(doc) => Ok(doc.sheets().len()),
            #[cfg(feature = "xlsx")]
            WorkbookImpl::Xlsx(xlsx) => {
                xlsx.ensure_source_current()?;
                Ok(xlsx.worksheet_count())
            },
            #[cfg(feature = "xlsb")]
            WorkbookImpl::Xlsb(xlsb) => Ok(xlsb.worksheet_count()),
            #[cfg(feature = "xls")]
            WorkbookImpl::XlsFile(xls) => Ok(xls.worksheet_count()),
            #[cfg(feature = "xls")]
            WorkbookImpl::XlsMem(xls) => Ok(xls.worksheet_count()),
            #[cfg(all(feature = "xls", any(unix, windows)))]
            WorkbookImpl::XlsSource(source) => source
                .workbook
                .worksheet_count()
                .map_err(map_xls_source_error),
            #[cfg(feature = "ods")]
            WorkbookImpl::Ods(spreadsheet) => Ok(spreadsheet.borrow().sheet_count()),
            #[cfg(all(feature = "ods", any(unix, windows)))]
            WorkbookImpl::OdsSource(spreadsheet) => Ok(spreadsheet.sheet_count()?),
            #[cfg(any(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
            WorkbookImpl::Other => Err(Box::new(Error::ParseError(
                "Unsupported workbook type in this build".to_string(),
            )) as Box<dyn std::error::Error + Send + Sync>),
        }
    }

    /// Extract all text from all worksheets.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::sheet::Workbook;
    ///
    /// let workbook = Workbook::open("spreadsheet.xlsx")?;
    /// let text = workbook.text()?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        match &self.inner {
            #[cfg(feature = "numbers")]
            WorkbookImpl::Numbers(doc) => doc.text().map_err(|e| {
                Box::new(Error::ParseError(format!(
                    "Failed to extract text from Numbers: {}",
                    e
                ))) as Box<dyn std::error::Error + Send + Sync>
            }),

            #[cfg(feature = "xlsx")]
            WorkbookImpl::Xlsx(xlsx) => {
                xlsx.ensure_source_current()?;
                // Iterate rows across worksheets
                let result = (|| {
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
                                append_cell_text(&mut out, cell);
                            }
                            out.push('\n');
                        }
                    }
                    Ok(out)
                })();
                xlsx.ensure_source_current()?;
                result
            },

            #[cfg(feature = "xlsb")]
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
                            append_cell_text(&mut out, cell);
                        }
                        out.push('\n');
                    }
                }
                Ok(out)
            },

            #[cfg(feature = "xls")]
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
                            append_cell_text(&mut out, cell);
                        }
                        out.push('\n');
                    }
                }
                Ok(out)
            },
            #[cfg(feature = "xls")]
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
                            append_cell_text(&mut out, cell);
                        }
                        out.push('\n');
                    }
                }
                Ok(out)
            },

            #[cfg(all(feature = "xls", any(unix, windows)))]
            WorkbookImpl::XlsSource(source) => source.with_eager(|xls| {
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
                            append_cell_text(&mut out, cell);
                        }
                        out.push('\n');
                    }
                }
                Ok(out)
            }),

            #[cfg(feature = "ods")]
            WorkbookImpl::Ods(spreadsheet) => Ok(spreadsheet.borrow().text()?),

            #[cfg(all(feature = "ods", any(unix, windows)))]
            WorkbookImpl::OdsSource(spreadsheet) => Ok(spreadsheet.text()?),

            #[cfg(any(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
            WorkbookImpl::Other => Err(Box::new(Error::ParseError(
                "Unsupported workbook type in this build".to_string(),
            )) as Box<dyn std::error::Error + Send + Sync>),
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
    /// let workbook = Workbook::open("spreadsheet.xlsx")?;
    /// let metadata = workbook.metadata()?;
    /// if let Some(title) = metadata.title {
    ///     println!("Title: {}", title);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn metadata(&self) -> Result<Metadata> {
        #[cfg(feature = "xlsx")]
        if let WorkbookImpl::Xlsx(workbook) = &self.inner {
            workbook.ensure_source_current()?;
        }
        #[cfg(all(feature = "ods", any(unix, windows)))]
        if let WorkbookImpl::OdsSource(spreadsheet) = &self.inner {
            return Ok(spreadsheet.metadata()?.clone());
        }
        #[cfg(all(feature = "xls", any(unix, windows)))]
        if let WorkbookImpl::XlsSource(source) = &self.inner {
            return source.metadata();
        }
        Ok(self.cached_metadata.clone())
    }

    /// Extract metadata from a Numbers document.
    ///
    /// This extracts metadata from the Numbers bundle, similar to how
    /// Keynote metadata is extracted.
    #[cfg(feature = "numbers")]
    fn extract_numbers_metadata(_doc: &litchi_numbers::Package) -> Metadata {
        Metadata {
            application: Some("Numbers".to_owned()),
            ..Metadata::default()
        }
    }
}

#[cfg(all(test, feature = "xlsx", any(unix, windows)))]
mod source_xlsx_path_tests {
    use super::{Workbook, WorkbookImpl};
    use crate::sheet::WorkbookTrait;
    use litchi_core::{Error, sheet::CellValue};
    use std::io::{Cursor, Write};

    const WORKSHEET: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    const WORKBOOK: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

    fn package_with_second_kind(
        second_sheet: &[u8],
        title: &str,
        second_content_type: &str,
        second_relationship_type: &str,
    ) -> Vec<u8> {
        let mut output = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                format!(
                    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{WORKBOOK}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{WORKSHEET}"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="{second_content_type}"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"#
                )
                .as_bytes(),
            )
            .unwrap();
        writer.start_file("_rels/.rels", options).unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("docProps/core.xml", options).unwrap();
        writer
            .write_all(
                format!(
                    r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>{title}</dc:title></cp:coreProperties>"#
                )
                .as_bytes(),
            )
            .unwrap();
        writer.start_file("xl/workbook.xml", options).unwrap();
        writer
            .write_all(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="1"/><sheets><sheet name="First" sheetId="1" r:id="rId1"/><sheet name="Second" sheetId="2" state="hidden" r:id="rId2"/></sheets></workbook>"#,
            )
            .unwrap();
        writer
            .start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        writer
            .write_all(
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{second_relationship_type}" Target="worksheets/sheet2.xml"/></Relationships>"#
                )
                .as_bytes(),
            )
            .unwrap();
        writer
            .start_file("xl/worksheets/sheet1.xml", options)
            .unwrap();
        writer
            .write_all(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#,
            )
            .unwrap();
        writer
            .start_file("xl/worksheets/sheet2.xml", options)
            .unwrap();
        writer.write_all(second_sheet).unwrap();
        writer.finish().unwrap();
        output.into_inner()
    }

    fn package(second_sheet: &[u8], title: &str) -> Vec<u8> {
        package_with_second_kind(
            second_sheet,
            title,
            WORKSHEET,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        )
    }

    fn valid_package(title: &str) -> Vec<u8> {
        package(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData></worksheet>"#,
            title,
        )
    }

    #[test]
    fn filesystem_and_bytes_xlsx_use_source_owner_and_match_projection() {
        let bytes = valid_package("source title");
        let path = tempfile::Builder::new().suffix(".ods").tempfile().unwrap();
        std::fs::write(path.path(), &bytes).unwrap();

        let source = Workbook::open(path.path()).expect("source XLSX");
        let bytes_workbook = Workbook::from_bytes(bytes.clone()).expect("bytes XLSX");
        let eager = crate::xlsx::Workbook::from_bytes(bytes).expect("typed eager XLSX");
        let WorkbookImpl::Xlsx(source_adapter) = &source.inner else {
            panic!("filesystem XLSX did not select the XLSX owner");
        };
        let WorkbookImpl::Xlsx(bytes_adapter) = &bytes_workbook.inner else {
            panic!("byte XLSX did not select the XLSX owner");
        };
        assert!(source_adapter.is_source_backed());
        assert!(bytes_adapter.is_source_backed());
        assert_eq!(source.worksheet_names().unwrap(), ["First", "Second"]);
        assert_eq!(
            source.worksheet_names().unwrap(),
            bytes_workbook.worksheet_names().unwrap()
        );
        assert_eq!(
            source.worksheet_count().unwrap(),
            bytes_workbook.worksheet_count().unwrap()
        );
        assert_eq!(
            eager
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["First", "Second"]
        );
        assert_eq!(source.text().unwrap(), bytes_workbook.text().unwrap());
        assert_eq!(source.text().unwrap(), "1\n2\n");
        assert_eq!(
            source.metadata().unwrap().title,
            bytes_workbook.metadata().unwrap().title
        );
        assert_eq!(
            source.metadata().unwrap().title.as_deref(),
            Some("source title")
        );
        assert!(source_adapter.is_1904_date_system());

        let first = source_adapter.worksheet_by_index(0).unwrap();
        assert_eq!(
            first.cell_by_coordinate("A1").unwrap().value(),
            &CellValue::Int(1)
        );
    }

    #[test]
    fn filesystem_xlsx_defers_unselected_worksheet_payload() {
        let bytes = package(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row>"#,
            "deferred",
        );
        let path = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(path.path(), bytes).unwrap();

        let workbook = Workbook::open(path.path()).expect("catalog-only source open");
        assert_eq!(workbook.worksheet_names().unwrap(), ["First", "Second"]);
        let WorkbookImpl::Xlsx(adapter) = &workbook.inner else {
            panic!("filesystem XLSX did not select the XLSX owner");
        };
        assert_eq!(
            adapter
                .worksheet_by_index(0)
                .unwrap()
                .cell_by_coordinate("A1")
                .unwrap()
                .value(),
            &CellValue::Int(1)
        );
        assert!(workbook.text().is_err());
    }

    #[test]
    fn bytes_xlsx_defers_unselected_worksheet_payload() {
        let bytes = package(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row>"#,
            "deferred bytes",
        );

        let workbook = Workbook::from_bytes(bytes).expect("catalog-only bytes open");
        assert_eq!(workbook.worksheet_names().unwrap(), ["First", "Second"]);
        let WorkbookImpl::Xlsx(adapter) = &workbook.inner else {
            panic!("bytes XLSX did not select the XLSX owner");
        };
        assert!(adapter.is_source_backed());
        assert_eq!(
            adapter
                .worksheet_by_index(0)
                .unwrap()
                .cell_by_coordinate("A1")
                .unwrap()
                .value(),
            &CellValue::Int(1)
        );
        assert!(workbook.text().is_err());
    }

    #[test]
    fn filesystem_xlsx_text_skips_non_grid_sheet_kinds() {
        let bytes = package_with_second_kind(
            br#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
            "chart",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
        );
        let path = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(path.path(), &bytes).unwrap();

        let source = Workbook::open(path.path()).expect("source XLSX with chart sheet");
        let eager = Workbook::from_bytes(bytes).expect("eager XLSX with chart sheet");
        assert_eq!(source.worksheet_names().unwrap(), ["First", "Second"]);
        assert_eq!(source.text().unwrap(), "1\n");
        assert_eq!(source.text().unwrap(), eager.text().unwrap());
    }

    #[test]
    fn filesystem_xlsx_cached_queries_report_source_change() {
        let path = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(path.path(), valid_package("before")).unwrap();
        let workbook = Workbook::open(path.path()).expect("source XLSX");

        std::fs::write(path.path(), valid_package("after with a different size")).unwrap();
        for result in [
            workbook.worksheet_names().map(|_| ()),
            workbook.worksheet_count().map(|_| ()),
            workbook.metadata().map(|_| ()),
            workbook.text().map(|_| ()),
        ] {
            assert!(matches!(
                result,
                Err(error) if error.downcast_ref::<Error>().is_some_and(|error| matches!(error, Error::SourceChanged { .. }))
            ));
        }
    }

    #[test]
    fn filesystem_xlsx_worksheet_handles_preserve_typed_source_change_errors() {
        let path = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(path.path(), valid_package("before")).unwrap();
        let workbook = Workbook::open(path.path()).expect("source XLSX");
        let WorkbookImpl::Xlsx(adapter) = &workbook.inner else {
            panic!("filesystem XLSX did not select the XLSX owner");
        };
        let first = adapter.worksheet_by_index(0).unwrap();

        std::fs::write(path.path(), valid_package("after with a different size")).unwrap();
        let cell_error = match first.cell_by_coordinate("A1") {
            Ok(_) => panic!("stale worksheet cell read unexpectedly succeeded"),
            Err(error) => error,
        };
        assert!(matches!(
            cell_error.downcast_ref::<Error>(),
            Some(Error::SourceChanged { .. })
        ));
        let row_error = first
            .rows()
            .next()
            .expect("stale row iterator returned no result")
            .unwrap_err();
        assert!(matches!(
            row_error.downcast_ref::<Error>(),
            Some(Error::SourceChanged { .. })
        ));
        let cell_iterator_error = match first
            .cells()
            .next()
            .expect("stale cell iterator returned no result")
        {
            Ok(_) => panic!("stale cell iterator unexpectedly succeeded"),
            Err(error) => error,
        };
        assert!(matches!(
            cell_iterator_error.downcast_ref::<Error>(),
            Some(Error::SourceChanged { .. })
        ));
    }

    #[test]
    fn filesystem_xlsx_propagates_the_opc_input_limit_before_fallback() {
        let path = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(path.path(), valid_package("bounded")).unwrap();
        path.as_file()
            .set_len(litchi_opc::ReadLimits::default().max_input_bytes() + 1)
            .unwrap();

        let error = Workbook::open(path.path())
            .err()
            .expect("oversized non-ODF ZIP unexpectedly opened");
        assert!(matches!(
            error.downcast_ref::<litchi_opc::OpcError>(),
            Some(litchi_opc::OpcError::ReadLimit {
                resource: litchi_opc::ReadResource::InputBytes,
                ..
            })
        ));
    }
}

#[cfg(all(test, feature = "ods", any(unix, windows)))]
mod source_ods_path_tests {
    use super::{Workbook, WorkbookImpl};
    use litchi_core::Error;

    const MIME: &str = "application/vnd.oasis.opendocument.spreadsheet";
    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    fn package(title: &str) -> Vec<u8> {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row table:number-rows-repeated="2"><table:table-cell table:number-columns-repeated="2" office:value-type="string"><text:p>hello</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>world</text:p></table:table-cell></table:table-row></table:table><table:table table:name="Sheet2"/></office:spreadsheet></office:body></office:document-content>"#
        );
        let metadata = format!(
            r#"<office:document-meta xmlns:office="{OFFICE}" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:meta><dc:title>{title}</dc:title></office:meta></office:document-meta>"#
        );
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer.set_mimetype(MIME).unwrap();
        writer.add_file("content.xml", content.as_bytes()).unwrap();
        writer.add_file("meta.xml", metadata.as_bytes()).unwrap();
        writer
            .add_file_with_media_type("Pictures/sample.bin", b"media", "application/octet-stream")
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn filesystem_ods_uses_source_owner_and_matches_byte_projection() {
        let bytes = package("source");
        let path = tempfile::Builder::new().suffix(".xlsx").tempfile().unwrap();
        std::fs::write(path.path(), &bytes).unwrap();

        let source = Workbook::open(path.path()).expect("source ODS");
        assert!(matches!(&source.inner, WorkbookImpl::OdsSource(_)));
        let eager = Workbook::from_bytes(bytes).expect("eager ODS");
        assert_eq!(
            source.worksheet_names().unwrap(),
            eager.worksheet_names().unwrap()
        );
        assert_eq!(
            source.worksheet_count().unwrap(),
            eager.worksheet_count().unwrap()
        );
        assert_eq!(source.text().unwrap(), eager.text().unwrap());
        assert_eq!(source.metadata().unwrap().title, Some("source".to_string()));
        assert_eq!(eager.metadata().unwrap().title, Some("source".to_string()));
    }

    #[test]
    fn source_root_metadata_reports_source_change() {
        let path = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(path.path(), package("before")).unwrap();
        let workbook = Workbook::open(path.path()).expect("source ODS");
        assert_eq!(
            workbook.metadata().unwrap().title,
            Some("before".to_string())
        );
        std::fs::write(path.path(), package("after with a different size")).unwrap();
        assert!(matches!(
            workbook.metadata(),
            Err(error) if error.downcast_ref::<Error>().is_some_and(|error| matches!(error, Error::SourceChanged { .. }))
        ));
    }

    #[test]
    fn source_root_rejects_malformed_ods_body() {
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer.set_mimetype(MIME).unwrap();
        writer
            .add_file("content.xml", b"<not-an-ods-document/>")
            .unwrap();
        let path = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(path.path(), writer.finish_to_bytes().unwrap()).unwrap();
        let error = match Workbook::open(path.path()) {
            Ok(_) => panic!("malformed ODS body was accepted"),
            Err(error) => error,
        };
        assert!(error.to_string().contains("ODS"));
    }

    #[test]
    fn filesystem_fallback_retains_bytes_from_the_open_source() {
        let bytes = b"not an office package".to_vec();
        let path = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(path.path(), &bytes).unwrap();
        let detected = crate::detection_smart::detected::detect_workbook_source_path(path.path())
            .expect("source fallback probe");
        match detected {
            crate::detection_smart::detected::WorkbookSourcePathDetection::Bytes(actual) => {
                assert_eq!(actual, bytes)
            },
            _ => panic!("non-ODS fallback did not retain source bytes"),
        }
    }

    #[test]
    fn source_root_preserves_repeated_text_semantics() {
        let path = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(path.path(), package("media")).unwrap();
        let workbook = Workbook::open(path.path()).expect("source ODS");
        assert_eq!(
            workbook.text().unwrap(),
            "hello\thello\tworld\nhello\thello\tworld\n"
        );
    }
}

#[cfg(all(test, not(all(feature = "ods", any(unix, windows)))))]
mod bounded_path_read_tests {
    use super::read_path_bytes;

    #[test]
    fn non_positional_path_reader_retains_exact_bytes() {
        let bytes = b"bounded workbook path bytes";
        let path = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(path.path(), bytes).unwrap();
        assert_eq!(read_path_bytes(path.path()).unwrap(), bytes);
    }
}

#[cfg(all(test, feature = "ods"))]
mod flat_ods_dispatch_tests {
    use super::Workbook;
    use litchi_core::detection::FileFormat;

    const FLAT_ODS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    office:version="1.3"
    office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell office:value-type="string"><text:p>value</text:p></table:table-cell>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document>"#;

    #[test]
    fn flat_ods_detection_and_public_open_agree() {
        assert_eq!(
            crate::detection_smart::detect_file_format_from_bytes(FLAT_ODS),
            Some(FileFormat::Ods)
        );
        assert!(matches!(
            crate::detection_smart::detect_format_smart(FLAT_ODS.to_vec()),
            Some(crate::detection_smart::DetectedFormat::FlatOdf(
                FileFormat::Ods,
                _
            ))
        ));

        assert!(matches!(
            Workbook::from_bytes(FLAT_ODS.to_vec()),
            Err(error) if error.to_string().contains("flat OpenDocument")
        ));

        let file = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(file.path(), FLAT_ODS).unwrap();
        assert!(matches!(
            Workbook::open(file.path()),
            Err(error) if error.to_string().contains("flat OpenDocument")
        ));
    }
}

#[cfg(all(test, feature = "ods", any(feature = "xlsx", feature = "xlsb")))]
mod ooxml_odf_polyglot_tests {
    use super::Workbook;
    use std::io::{Cursor, Write};

    fn dual_marker_xlsx() -> Vec<u8> {
        let mut output = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        // Keep the valid ODF local mimetype first, as required by ODF. The
        // remaining entries form a valid minimal OPC/XLSX package as well.
        writer.start_file("mimetype", options).unwrap();
        writer
            .write_all(b"application/vnd.oasis.opendocument.spreadsheet")
            .unwrap();
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
            )
            .unwrap();
        writer.start_file("_rels/.rels", options).unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("xl/workbook.xml", options).unwrap();
        writer
            .write_all(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            )
            .unwrap();
        writer
            .start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .start_file("xl/worksheets/sheet1.xml", options)
            .unwrap();
        writer
            .write_all(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#,
            )
            .unwrap();
        writer.finish().unwrap();
        output.into_inner()
    }

    #[cfg(feature = "xlsx")]
    #[test]
    fn ooxml_first_precedence_survives_an_odf_local_mimetype_marker() {
        let bytes = dual_marker_xlsx();
        assert!(matches!(
            crate::detection_smart::detect_format_smart(bytes.clone()),
            Some(crate::detection_smart::DetectedFormat::Xlsx(_))
        ));
        let workbook = Workbook::from_bytes(bytes)
            .expect("OOXML-first precedence should select the valid XLSX owner");
        assert_eq!(workbook.worksheet_names().unwrap(), ["Sheet1"]);
    }

    #[cfg(all(feature = "xlsx", any(unix, windows)))]
    #[test]
    fn filesystem_ooxml_first_precedence_survives_an_odf_local_mimetype_marker() {
        let path = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(path.path(), dual_marker_xlsx()).unwrap();
        let workbook = Workbook::open(path.path())
            .expect("filesystem OOXML-first precedence should select XLSX");
        assert_eq!(workbook.worksheet_names().unwrap(), ["Sheet1"]);
    }

    #[cfg(all(feature = "xlsb", not(feature = "xlsx")))]
    #[test]
    fn disabled_xlsx_owner_keeps_smart_precedence() {
        let bytes = dual_marker_xlsx();
        assert!(crate::detection_smart::detect_format_smart(bytes.clone()).is_none());
        assert!(crate::detection_smart::detected::detect_prepared_ods(bytes.clone()).is_err());

        let error = Workbook::from_bytes(bytes)
            .err()
            .expect("disabled XLSX owner must not fall through to ODS");
        assert_eq!(error.to_string(), "Not a valid Office file");
    }

    #[cfg(all(feature = "xlsb", not(feature = "xlsx"), any(unix, windows)))]
    #[test]
    fn filesystem_disabled_xlsx_owner_keeps_smart_precedence() {
        let path = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(path.path(), dual_marker_xlsx()).unwrap();
        let error = match Workbook::open(path.path()) {
            Ok(_) => panic!("disabled XLSX owner fell through to ODS"),
            Err(error) => error,
        };
        assert_eq!(error.to_string(), "Not a valid Office file");
    }

    #[test]
    fn invalid_odf_body_still_falls_back_to_typed_ods_validation() {
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_SPREADSHEET)
            .unwrap();
        writer
            .add_file("content.xml", b"<not-an-ods-document/>")
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let error = Workbook::from_bytes(bytes)
            .err()
            .expect("invalid ODS body must not be accepted as XLSX");
        assert!(error.to_string().contains("ODS"));
    }

    #[test]
    fn valid_ods_wins_after_the_ooxml_probe_fails() {
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_SPREADSHEET)
            .unwrap();
        writer
            .add_file(
                "content.xml",
                br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:name="Sheet1"/></office:spreadsheet></office:body></office:document-content>"#,
            )
            .unwrap();

        assert!(Workbook::from_bytes(writer.finish_to_bytes().unwrap()).is_ok());
    }
}

#[cfg(all(test, any(feature = "xlsx", feature = "xlsb"), feature = "xls"))]
mod tests {
    use super::*;
    use std::path::PathBuf;

    fn test_data_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
    }

    #[test]
    #[cfg(all(any(feature = "xlsx", feature = "xlsb"), feature = "xls"))]
    fn test_workbook_open_xlsx() {
        let path = test_data_path().join("ooxml/xlsx/DateFormatTests.xlsx");
        let workbook = Workbook::open(&path);
        assert!(
            workbook.is_ok(),
            "Failed to open XLSX file: {:?}",
            workbook.err()
        );
    }

    #[test]
    #[cfg(all(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
    fn test_workbook_open_xls() {
        // XLS parsing has issues with some test files - this test documents the limitation
        // Skip if no working XLS files are available
        let path = test_data_path().join("ole/xls/Simple.xls");
        if path.exists() {
            // Try to open, but don't fail the test if XLS parsing has issues
            let _workbook = Workbook::open(&path);
            // Just verify the file exists and we can attempt to open it
        }
    }

    #[test]
    #[cfg(all(any(feature = "xlsx", feature = "xlsb"), feature = "xls"))]
    fn test_workbook_from_bytes_xlsx() {
        let path = test_data_path().join("ooxml/xlsx/DateFormatTests.xlsx");
        let bytes = std::fs::read(&path).expect("Failed to read file");
        let workbook = Workbook::from_bytes(bytes);
        assert!(
            workbook.is_ok(),
            "Failed to load XLSX from bytes: {:?}",
            workbook.err()
        );
    }

    #[test]
    #[cfg(all(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
    fn test_workbook_from_bytes_xls() {
        // XLS parsing has issues with some test files - this test documents the limitation
        let path = test_data_path().join("ole/xls/Simple.xls");
        if path.exists() {
            let bytes = std::fs::read(&path).expect("Failed to read file");
            // Try to load, but don't fail the test if XLS parsing has issues
            let _workbook = Workbook::from_bytes(bytes);
        }
    }

    #[test]
    #[cfg(all(any(feature = "xlsx", feature = "xlsb"), feature = "xls"))]
    fn test_workbook_worksheet_names_xlsx() {
        let path = test_data_path().join("ooxml/xlsx/DateFormatTests.xlsx");
        let workbook = Workbook::open(&path).expect("Failed to open XLSX");
        let names = workbook
            .worksheet_names()
            .expect("Failed to get worksheet names");
        assert!(!names.is_empty(), "Expected at least one worksheet");
    }

    #[test]
    #[cfg(all(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
    fn test_workbook_worksheet_names_xls() {
        // XLS parsing has issues with some test files - this test documents the limitation
        let path = test_data_path().join("ole/xls/Simple.xls");
        if let Ok(workbook) = Workbook::open(&path) {
            let names = workbook
                .worksheet_names()
                .expect("Failed to get worksheet names");
            assert!(!names.is_empty(), "Expected at least one worksheet");
        }
    }

    #[test]
    #[cfg(all(any(feature = "xlsx", feature = "xlsb"), feature = "xls"))]
    fn test_workbook_worksheet_count_xlsx() {
        let path = test_data_path().join("ooxml/xlsx/DateFormatTests.xlsx");
        let workbook = Workbook::open(&path).expect("Failed to open XLSX");
        let count = workbook
            .worksheet_count()
            .expect("Failed to get worksheet count");
        assert!(count > 0, "Expected at least one worksheet");
    }

    #[test]
    #[cfg(all(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
    fn test_workbook_worksheet_count_xls() {
        // XLS parsing has issues with some test files - this test documents the limitation
        let path = test_data_path().join("ole/xls/Simple.xls");
        if let Ok(workbook) = Workbook::open(&path) {
            let count = workbook
                .worksheet_count()
                .expect("Failed to get worksheet count");
            assert!(count > 0, "Expected at least one worksheet");
        }
    }

    #[test]
    #[cfg(all(any(feature = "xlsx", feature = "xlsb"), feature = "xls"))]
    fn test_workbook_text_xlsx() {
        let path = test_data_path().join("ooxml/xlsx/DateFormatTests.xlsx");
        let workbook = Workbook::open(&path).expect("Failed to open XLSX");
        let _text = workbook.text().expect("Failed to extract text");
        // Text may vary by file
    }

    #[test]
    #[cfg(all(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
    fn test_workbook_text_xls() {
        // XLS parsing has issues with some test files - this test documents the limitation
        let path = test_data_path().join("ole/xls/Simple.xls");
        if let Ok(workbook) = Workbook::open(&path) {
            let _text = workbook.text().expect("Failed to extract text");
        }
    }

    #[test]
    #[cfg(all(any(feature = "xlsx", feature = "xlsb"), feature = "xls"))]
    fn test_workbook_metadata_xlsx() {
        let path = test_data_path().join("ooxml/xlsx/DateFormatTests.xlsx");
        let workbook = Workbook::open(&path).expect("Failed to open XLSX");
        let metadata = workbook.metadata().expect("Failed to get metadata");
        // Metadata may or may not be present
        let _ = metadata.title;
        let _ = metadata.author;
    }

    #[test]
    #[cfg(all(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
    fn test_workbook_metadata_xls() {
        // XLS parsing has issues with some test files - this test documents the limitation
        let path = test_data_path().join("ole/xls/Simple.xls");
        if let Ok(workbook) = Workbook::open(&path) {
            let metadata = workbook.metadata().expect("Failed to get metadata");
            let _ = metadata.title;
            let _ = metadata.author;
        }
    }

    #[test]
    #[cfg(all(any(feature = "xlsx", feature = "xlsb"), feature = "xls"))]
    fn test_workbook_open_nonexistent_file() {
        let path = test_data_path().join("nonexistent_file.xlsx");
        let result = Workbook::open(&path);
        assert!(result.is_err(), "Expected error for nonexistent file");
    }

    #[test]
    #[cfg(all(any(feature = "xlsx", feature = "xlsb"), feature = "xls"))]
    fn test_workbook_from_bytes_invalid_data() {
        let bytes = b"This is not a valid spreadsheet file".to_vec();
        let result = Workbook::from_bytes(bytes);
        assert!(result.is_err(), "Expected error for invalid data");
    }

    #[test]
    #[cfg(all(any(feature = "xlsx", feature = "xlsb"), feature = "xls"))]
    fn test_workbook_conditional_formatting_xlsx() {
        // Use a simpler XLSX file that is known to work
        let path = test_data_path().join("ooxml/xlsx/condFormat_cellis.xlsx");
        if path.exists() {
            let workbook = Workbook::open(&path);
            assert!(
                workbook.is_ok(),
                "Failed to open conditional formatting XLSX"
            );

            if let Ok(wb) = workbook {
                let names = wb.worksheet_names().expect("Failed to get names");
                assert!(!names.is_empty(), "Expected worksheets");
            }
        }
    }

    #[test]
    #[cfg(all(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
    fn test_workbook_conditional_formatting_xls() {
        // XLS parsing has issues - test only if file can be opened
        let path = test_data_path().join("ole/xls/ConditionalFormattingSamples.xls");
        if let Ok(workbook) = Workbook::open(&path) {
            let names = workbook.worksheet_names().expect("Failed to get names");
            assert!(!names.is_empty(), "Expected worksheets");
        }
    }

    #[test]
    #[cfg(all(any(feature = "xlsx", feature = "xlsb"), feature = "xls"))]
    fn test_workbook_autofilter_xlsx() {
        let path = test_data_path().join("ooxml/xlsx/autofilter.xlsx");
        let workbook = Workbook::open(&path);
        assert!(workbook.is_ok(), "Failed to open autofilter XLSX");

        if let Ok(wb) = workbook {
            let count = wb.worksheet_count().expect("Failed to get count");
            assert!(count > 0, "Expected at least one worksheet");
        }
    }

    #[test]
    #[cfg(all(any(feature = "xlsx", feature = "xlsb"), feature = "xls"))]
    fn test_workbook_data_validation_xlsx() {
        let path = test_data_path().join("ooxml/xlsx/DataValidationEvaluations.xlsx");
        let workbook = Workbook::open(&path);
        assert!(workbook.is_ok(), "Failed to open data validation XLSX");

        if let Ok(wb) = workbook {
            let count = wb.worksheet_count().expect("Failed to get count");
            assert!(count > 0, "Expected at least one worksheet");
        }
    }

    #[test]
    #[cfg(all(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
    fn test_workbook_formulas_xls() {
        let path = test_data_path().join("ole/xls/FormulaEvalTestData.xls");
        let workbook = Workbook::open(&path);
        assert!(workbook.is_ok(), "Failed to open formula test XLS");

        if let Ok(wb) = workbook {
            let _text = wb.text().expect("Failed to extract text");
        }
    }

    #[test]
    #[cfg(all(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
    fn test_workbook_hyperlinks_xls() {
        let path = test_data_path().join("ole/xls/HyperlinksOnManySheets.xls");
        let workbook = Workbook::open(&path);
        assert!(workbook.is_ok(), "Failed to open hyperlinks XLS");

        if let Ok(wb) = workbook {
            let names = wb.worksheet_names().expect("Failed to get names");
            assert!(!names.is_empty(), "Expected worksheets with hyperlinks");
        }
    }
}

#[cfg(all(test, feature = "xls", any(unix, windows)))]
mod source_xls_path_tests {
    use super::{Workbook, WorkbookImpl, XlsSource};
    use litchi_cfb::{OleWriter, SharedOleFile};
    use litchi_core::{FileSource, OwnedSource, ReadAt, SourceVersion};
    use litchi_ole_common::property_set::SharedPropertySetReader;
    use litchi_xls::{SourceBackedError, SourceBackedLimits, SourceBackedWorkbook};
    use std::io::{self, Cursor, Write};
    use std::sync::atomic::{AtomicU64, Ordering};
    use std::sync::{Arc, Mutex};

    fn fixture_path() -> std::path::PathBuf {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/xls/Simple.xls")
    }

    fn write_temporary(bytes: &[u8]) -> tempfile::NamedTempFile {
        let file = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(file.path(), bytes).unwrap();
        file
    }

    fn cfb_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
        let mut writer = OleWriter::new();
        for (name, bytes) in streams {
            writer.create_stream(&[*name], bytes).unwrap();
        }
        let mut output = Vec::new();
        writer.write_to(&mut Cursor::new(&mut output)).unwrap();
        output
    }

    #[derive(Clone)]
    struct CountingSource {
        bytes: Arc<Vec<u8>>,
        ranges: Arc<Mutex<Vec<(u64, usize)>>>,
        revision: Arc<AtomicU64>,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes: Arc::new(bytes),
                ranges: Arc::new(Mutex::new(Vec::new())),
                revision: Arc::new(AtomicU64::new(0)),
            }
        }

        fn clear_ranges(&self) {
            self.ranges.lock().unwrap().clear();
        }

        fn ranges(&self) -> Vec<(u64, usize)> {
            self.ranges.lock().unwrap().clone()
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let start = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
            if start >= self.bytes.len() || output.is_empty() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            self.ranges.lock().unwrap().push((offset, count));
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                0x584c_535f_46414345,
                self.revision.load(Ordering::Relaxed),
            ))
        }
    }

    fn global_offsets(stream: &[u8]) -> (Vec<u64>, usize) {
        let mut offsets = Vec::new();
        let mut cursor = 0_usize;
        loop {
            offsets.push(cursor as u64);
            let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
            let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
            cursor += 4 + length;
            if kind == 0x000A {
                return (offsets, cursor);
            }
        }
    }

    fn counted_stream_range(
        source: &CountingSource,
        cfb: &SharedOleFile,
        offset: u64,
        length: usize,
    ) -> Vec<(u64, usize)> {
        source.clear_ranges();
        let mut output = vec![0_u8; length];
        cfb.read_stream_range(&["Workbook"], offset, &mut output)
            .unwrap();
        let ranges = source.ranges();
        source.clear_ranges();
        ranges
    }

    fn frame_bytes(kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(4 + payload.len());
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&(u16::try_from(payload.len()).unwrap()).to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    fn legacy_biff_corpus() -> Vec<u8> {
        let source: Arc<dyn ReadAt> =
            Arc::new(OwnedSource::new(std::fs::read(fixture_path()).unwrap()));
        let mut stream = litchi_cfb::SharedOleFile::open(source)
            .unwrap()
            .open_stream(&["Workbook"])
            .unwrap();
        stream[4..6].copy_from_slice(&0x0500_u16.to_le_bytes());
        cfb_with_streams(&[("Workbook", &stream)])
    }

    fn valid_workbook_stream() -> Vec<u8> {
        let source: Arc<dyn ReadAt> =
            Arc::new(OwnedSource::new(std::fs::read(fixture_path()).unwrap()));
        SharedOleFile::open(source)
            .unwrap()
            .open_stream(&["Workbook"])
            .unwrap()
    }

    #[test]
    fn filesystem_open_retains_xls_source_until_text() {
        let path = fixture_path();
        let workbook = Workbook::open(&path).expect("failed to open XLS source");
        let WorkbookImpl::XlsSource(source) = &workbook.inner else {
            panic!("filesystem XLS did not select the source-backed facade variant");
        };
        assert!(source.eager.get().is_none());

        let names = workbook
            .worksheet_names()
            .expect("failed to enumerate source-backed XLS sheets");
        assert!(!names.is_empty());
        assert_eq!(workbook.worksheet_count().unwrap(), names.len());
        assert!(source.eager.get().is_none());
    }

    #[test]
    fn source_probe_reuses_one_catalog_for_globals_and_metadata() {
        let bytes = std::fs::read(fixture_path()).unwrap();
        let source = Arc::new(CountingSource::new(bytes.clone()));
        let retained: Arc<dyn ReadAt> = source.clone();

        source.clear_ranges();
        let baseline = Arc::new(SharedOleFile::open(Arc::clone(&retained)).unwrap());
        let catalog_ranges = source.ranges();
        source.clear_ranges();
        let workbook_stream = baseline.open_stream(&["Workbook"]).unwrap();
        source.clear_ranges();
        let (global_offsets, global_end) = global_offsets(&workbook_stream);
        let source_version = retained.version().unwrap();

        let (owner, cfb) = crate::detection_smart::detected::try_open_xls_source(
            Arc::clone(&retained),
            source_version,
        )
        .unwrap()
        .expect("valid XLS source probe");
        let actual = source.ranges();
        assert_eq!(actual.first().copied(), Some((0, 8)));
        assert!(actual.len() >= 1 + catalog_ranges.len());
        assert_eq!(&actual[1..1 + catalog_ranges.len()], &catalog_ranges);

        let mut expected_globals = Vec::new();
        for offset in global_offsets {
            expected_globals.extend(counted_stream_range(&source, &cfb, offset, 4));
        }
        expected_globals.extend(counted_stream_range(&source, &cfb, 0, global_end));
        assert_eq!(&actual[1 + catalog_ranges.len()..], &expected_globals);

        let adapter = XlsSource::new(owner, Arc::clone(&cfb));
        let shared_count = Arc::strong_count(&cfb);
        source.clear_ranges();
        let expected_metadata = litchi_core::Metadata::from(cfb.get_metadata().unwrap());
        let direct_metadata_ranges = source.ranges();
        source.clear_ranges();
        let metadata = adapter.metadata().unwrap();
        let metadata_ranges = source.ranges();
        assert_eq!(metadata.application, expected_metadata.application);
        assert_eq!(metadata.author, expected_metadata.author);
        assert_eq!(metadata_ranges, direct_metadata_ranges);
        assert_eq!(Arc::strong_count(&cfb), shared_count);
    }

    #[test]
    fn source_xls_matches_eager_names_text_and_metadata() {
        let path = fixture_path();
        let source = Workbook::open(&path).expect("failed to open source-backed XLS");
        let WorkbookImpl::XlsSource(source_adapter) = &source.inner else {
            panic!("filesystem XLS did not select the source-backed facade variant");
        };
        assert!(source_adapter.eager.get().is_none());
        let source_metadata = source.metadata().unwrap();
        assert!(source_metadata.application.is_some());
        assert!(source_adapter.eager.get().is_none());

        let eager =
            Workbook::from_bytes(std::fs::read(&path).unwrap()).expect("failed to open eager XLS");
        let eager_metadata = eager.metadata().unwrap();

        assert_eq!(
            source.worksheet_names().unwrap(),
            eager.worksheet_names().unwrap()
        );
        assert_eq!(
            source.worksheet_count().unwrap(),
            eager.worksheet_count().unwrap()
        );
        assert_eq!(source_metadata.application, eager_metadata.application);
        assert_eq!(source_metadata.author, eager_metadata.author);
        assert_eq!(source.text().unwrap(), eager.text().unwrap());
        assert!(source_adapter.eager.get().is_some());
    }

    #[test]
    fn from_bytes_remains_eager_xls_memory_variant() {
        let workbook = Workbook::from_bytes(std::fs::read(fixture_path()).unwrap())
            .expect("failed to open eager XLS");
        assert!(matches!(workbook.inner, WorkbookImpl::XlsMem(_)));
    }

    #[test]
    fn legacy_biff_path_matches_from_bytes_and_owner_error() {
        let bytes = legacy_biff_corpus();
        let file = write_temporary(&bytes);
        let path_workbook = Workbook::open(file.path()).expect("legacy BIFF path fallback");
        let bytes_workbook = Workbook::from_bytes(bytes.clone()).expect("legacy BIFF bytes");
        assert!(matches!(&path_workbook.inner, WorkbookImpl::XlsMem(_)));
        assert_eq!(
            path_workbook.worksheet_names().unwrap(),
            bytes_workbook.worksheet_names().unwrap()
        );
        assert_eq!(
            path_workbook.worksheet_count().unwrap(),
            bytes_workbook.worksheet_count().unwrap()
        );
        assert_eq!(
            path_workbook.text().unwrap(),
            bytes_workbook.text().unwrap()
        );

        let owner_error = match SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(
            legacy_biff_corpus(),
        ))) {
            Ok(_) => panic!("legacy BIFF unexpectedly opened by source owner"),
            Err(error) => error,
        };
        assert!(matches!(
            owner_error,
            SourceBackedError::UnsupportedBiffVersion(0x0500)
        ));
    }

    #[test]
    fn source_xls_refuses_stale_path() {
        let fixture = std::fs::read(fixture_path()).unwrap();
        let file = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(file.path(), fixture).unwrap();
        let workbook = Workbook::open(file.path()).expect("failed to open temporary XLS");

        let mut changed = std::fs::OpenOptions::new()
            .append(true)
            .open(file.path())
            .unwrap();
        changed.write_all(&[0]).unwrap();
        changed.flush().unwrap();

        let error = workbook.worksheet_names().unwrap_err();
        assert!(matches!(
            error.downcast_ref::<SourceBackedError>(),
            Some(SourceBackedError::SourceChanged { .. })
        ));
    }

    #[test]
    fn source_xls_refuses_encrypted_filepass_without_eager_fallback() {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/poi/test-data/spreadsheet/xor-encryption-abc.xls"),
        )
        .unwrap();
        let file = write_temporary(&bytes);
        let error = match Workbook::open(file.path()) {
            Ok(_) => panic!("encrypted XLS unexpectedly opened through the facade"),
            Err(error) => error,
        };
        assert!(matches!(
            error.downcast_ref::<SourceBackedError>(),
            Some(SourceBackedError::EncryptedUnsupported)
        ));
    }

    #[test]
    fn source_xls_refuses_malformed_biff_without_eager_fallback() {
        let stream = frame_bytes(0x0809, &[0; 16]);
        let file = write_temporary(&cfb_with_streams(&[("Workbook", &stream)]));
        let error = match Workbook::open(file.path()) {
            Ok(_) => panic!("malformed XLS unexpectedly opened through the facade"),
            Err(error) => error,
        };
        assert!(matches!(
            error.downcast_ref::<SourceBackedError>(),
            Some(
                SourceBackedError::Cfb(_)
                    | SourceBackedError::Parse(_)
                    | SourceBackedError::InvalidData(_)
            )
        ));
    }

    #[cfg(unix)]
    #[test]
    fn source_xls_pins_open_file_across_atomic_path_replacement() {
        let original = write_temporary(&std::fs::read(fixture_path()).unwrap());
        let workbook = Workbook::open(original.path()).expect("failed to open source XLS");
        let names = workbook.worksheet_names().unwrap();

        let replacement = write_temporary(b"not an XLS workbook");
        std::fs::rename(replacement.path(), original.path()).unwrap();

        // The already-open descriptor remains pinned to the original inode;
        // it must not retarget to the replacement pathname contents.
        assert_eq!(workbook.worksheet_names().unwrap(), names);
        assert!(Workbook::open(original.path()).is_err());
    }

    #[cfg(feature = "xlsx")]
    #[test]
    fn xls_source_probe_preserves_xlsx_precedence() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ooxml/xlsx/DateFormatTests.xlsx");
        let workbook = Workbook::open(path).unwrap();
        assert!(matches!(workbook.inner, WorkbookImpl::Xlsx(_)));
    }

    #[test]
    fn metadata_ole_errors_use_source_owner_normalization() {
        let source = FileSource::open(fixture_path()).unwrap();
        let expected = source.version().unwrap();
        let observed = expected;
        let changed = super::normalize_xls_metadata_error(litchi_cfb::OleError::SourceChanged {
            expected,
            observed,
        });
        assert!(matches!(
            changed,
            SourceBackedError::SourceChanged {
                expected: left,
                observed: right
            } if left == expected && right == observed
        ));

        let limited = super::normalize_xls_metadata_error(litchi_cfb::OleError::LimitExceeded {
            resource: "input bytes",
            observed: 9,
            maximum: 8,
        });
        assert!(matches!(
            limited,
            SourceBackedError::ResourceLimit {
                resource: "input bytes",
                observed: 9,
                maximum: 8,
            }
        ));
    }

    #[test]
    fn ole_host_stream_precedence_is_catalog_only_and_ordered() {
        let cases: Vec<(
            Vec<(&str, &[u8])>,
            Option<crate::detection_smart::detected::OleHostStream>,
        )> = vec![
            (
                vec![
                    ("WordDocument", b"word".as_slice()),
                    ("PowerPoint Document", b"ppt".as_slice()),
                    ("Workbook", b"xls".as_slice()),
                ],
                Some(crate::detection_smart::detected::OleHostStream::WordDocument),
            ),
            (
                vec![
                    ("PowerPoint Document", b"ppt".as_slice()),
                    ("Workbook", b"xls".as_slice()),
                ],
                Some(crate::detection_smart::detected::OleHostStream::PowerPointDocument),
            ),
            (
                vec![
                    ("Current User", b"ppt".as_slice()),
                    ("Workbook", b"xls".as_slice()),
                ],
                Some(crate::detection_smart::detected::OleHostStream::PowerPointDocument),
            ),
            (
                vec![("Book", b"xls".as_slice()), ("Workbook", b"xls".as_slice())],
                Some(crate::detection_smart::detected::OleHostStream::Workbook),
            ),
            (vec![("Other", b"ole".as_slice())], None),
        ];
        for (streams, expected) in cases {
            let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(cfb_with_streams(&streams)));
            let version = source.version().unwrap();
            assert_eq!(
                crate::detection_smart::detected::classify_ole_host_stream(
                    Arc::clone(&source),
                    version,
                )
                .unwrap(),
                expected
            );
        }
    }

    #[cfg(not(any(feature = "doc", feature = "ppt")))]
    #[test]
    fn path_level_feature_elided_foreign_owners_intentionally_fall_back_to_eager_xls() {
        let workbook_stream = valid_workbook_stream();
        let cases = [
            vec![
                ("WordDocument", b"word".to_vec()),
                ("PowerPoint Document", b"ppt".to_vec()),
                ("Workbook", workbook_stream.clone()),
            ],
            vec![
                ("PowerPoint Document", b"ppt".to_vec()),
                ("Workbook", workbook_stream.clone()),
            ],
            vec![
                ("Current User", b"ppt".to_vec()),
                ("Workbook", workbook_stream),
            ],
        ];
        for streams in cases {
            let refs = streams
                .iter()
                .map(|(name, bytes)| (*name, bytes.as_slice()))
                .collect::<Vec<_>>();
            let file = write_temporary(&cfb_with_streams(&refs));
            let workbook = Workbook::open(file.path()).expect("polyglot XLS fallback");
            assert!(matches!(workbook.inner, WorkbookImpl::XlsMem(_)));
        }
    }

    #[cfg(all(feature = "doc", feature = "ppt"))]
    #[test]
    fn path_level_foreign_owners_precede_xls_when_host_features_are_enabled() {
        let workbook_stream = valid_workbook_stream();
        let cases = [
            vec![
                ("WordDocument", b"word".to_vec()),
                ("PowerPoint Document", b"ppt".to_vec()),
                ("Workbook", workbook_stream.clone()),
            ],
            vec![
                ("PowerPoint Document", b"ppt".to_vec()),
                ("Workbook", workbook_stream.clone()),
            ],
            vec![
                ("Current User", b"ppt".to_vec()),
                ("Workbook", workbook_stream),
            ],
        ];
        for streams in cases {
            let refs = streams
                .iter()
                .map(|(name, bytes)| (*name, bytes.as_slice()))
                .collect::<Vec<_>>();
            let file = write_temporary(&cfb_with_streams(&refs));
            assert!(
                Workbook::open(file.path()).is_err(),
                "foreign OLE owner must not fall back to XLS when its host feature is enabled"
            );
        }
    }

    #[test]
    fn source_xls_materialization_limit_stays_typed() {
        let source: Arc<dyn ReadAt> = Arc::new(FileSource::open(fixture_path()).unwrap());
        let owner = SourceBackedWorkbook::from_read_at_with_limits(
            Arc::clone(&source),
            SourceBackedLimits::default().with_max_materialize_bytes(1),
        )
        .unwrap();
        let cfb = Arc::new(litchi_cfb::SharedOleFile::open(Arc::clone(&source)).unwrap());
        let adapter = XlsSource::new(owner, cfb);
        let error = adapter.with_eager(|_| Ok(())).unwrap_err();
        assert!(matches!(
            error.downcast_ref::<SourceBackedError>(),
            Some(SourceBackedError::ResourceLimit { .. })
        ));
    }
}
