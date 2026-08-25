//! Read-only source-backed access to existing BIFF8 XLS workbooks.
//!
//! Opening this owner validates the CFB directory, workbook globals, and
//! worksheet offset topology without materializing worksheet payload. A
//! selected cell query incrementally reads and parses only the selected
//! worksheet through its valid EOF. Convenience queries run sequentially;
//! callers that need cooperative cancellation can use the explicit
//! `*_with_execution` variants. Finite scan limits remain mandatory for
//! bounded work.

use crate::cell::Cell;
use crate::error::Error;
use crate::leniency::{Leniency, ToleranceLog};
use crate::number_format::{DateSystem, Formatting};
use crate::records::{BofRecord, BoundSheetRecord, CellRecord, Encoding, FormulaValue, SheetType};
use crate::{SheetKind, SheetVisibility, Workbook};
use litchi_biff::{Limits as BiffLimits, RecordRef, Records as BiffRecords};
use litchi_cfb::{
    OleError, SharedOleFile, SharedOleFileLimits, SharedOleStreamSession,
    SharedOleStreamSessionOutcome,
};
#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::sheet::Cell as CellTrait;
use litchi_core::{ExecutionContext, ExecutionError, ReadAt, SourceVersion};
use std::collections::HashSet;
use std::fmt;
use std::io::Cursor;
use std::path::Path;
use std::sync::Arc;

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000A;
const CODEPAGE: u16 = 0x0042;
const BOUND_SHEET: u16 = 0x0085;
const FILEPASS: u16 = 0x002F;
const SST: u16 = 0x00FC;
const CONTINUE: u16 = 0x003C;
const STRING: u16 = 0x0207;
const BIFF8: u16 = 0x0600;
const WORKBOOK_BOF_TYPE: u16 = 0x0005;
const WORKSHEET_BOF_TYPE: u16 = 0x0010;
const DEFAULT_CODEPAGE: u16 = 1252;
const DEFAULT_GLOBAL_BYTES: u64 = 128 * 1024 * 1024;
const DEFAULT_GLOBAL_RECORDS: usize = 1_000_000;
const DEFAULT_SST_ENTRIES: usize = 1_000_000;
const DEFAULT_WORKSHEET_BYTES: u64 = 128 * 1024 * 1024;
const DEFAULT_WORKSHEET_RECORDS: usize = 1_000_000;
const DEFAULT_SHEET_COUNT: usize = 4_096;
const DEFAULT_MATERIALIZE_BYTES: u64 = 128 * 1024 * 1024;
const MATERIALIZE_CHUNK_BYTES: usize = 64 * 1024;

/// Limits for one source-backed XLS owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceBackedLimits {
    /// Maximum physical CFB input size.
    pub max_input_bytes: u64,
    /// Maximum bytes copied by an eager materialization fallback.
    pub max_materialize_bytes: u64,
    /// Maximum bytes retained while parsing workbook globals.
    pub max_global_bytes: u64,
    /// Maximum BIFF records in workbook globals.
    pub max_global_records: usize,
    /// Maximum unique or total SST entries accepted.
    pub max_sst_entries: usize,
    /// Maximum logical bytes traversed for one worksheet query.
    pub max_worksheet_scan_bytes: u64,
    /// Maximum BIFF records parsed for one worksheet query.
    pub max_worksheet_scan_records: usize,
    /// Maximum number of `BoundSheet8` entries.
    pub max_sheet_count: usize,
}

impl Default for SourceBackedLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: SharedOleFileLimits::MAX_INPUT_BYTES,
            max_materialize_bytes: DEFAULT_MATERIALIZE_BYTES,
            max_global_bytes: DEFAULT_GLOBAL_BYTES,
            max_global_records: DEFAULT_GLOBAL_RECORDS,
            max_sst_entries: DEFAULT_SST_ENTRIES,
            max_worksheet_scan_bytes: DEFAULT_WORKSHEET_BYTES,
            max_worksheet_scan_records: DEFAULT_WORKSHEET_RECORDS,
            max_sheet_count: DEFAULT_SHEET_COUNT,
        }
    }
}

impl SourceBackedLimits {
    /// Creates defaults with an explicit CFB input ceiling.
    pub fn new(max_input_bytes: u64) -> std::result::Result<Self, SourceBackedError> {
        let limits = Self {
            max_input_bytes,
            ..Self::default()
        };
        limits.validate()?;
        Ok(limits)
    }

    /// Sets the CFB input ceiling.
    #[must_use]
    pub const fn with_max_input_bytes(mut self, value: u64) -> Self {
        self.max_input_bytes = value;
        self
    }

    /// Sets the eager-materialization byte ceiling.
    #[must_use]
    pub const fn with_max_materialize_bytes(mut self, value: u64) -> Self {
        self.max_materialize_bytes = value;
        self
    }

    /// Sets the workbook-global byte ceiling.
    #[must_use]
    pub const fn with_max_global_bytes(mut self, value: u64) -> Self {
        self.max_global_bytes = value;
        self
    }

    /// Sets the workbook-global record ceiling.
    #[must_use]
    pub const fn with_max_global_records(mut self, value: usize) -> Self {
        self.max_global_records = value;
        self
    }

    /// Sets the shared-string entry ceiling.
    #[must_use]
    pub const fn with_max_sst_entries(mut self, value: usize) -> Self {
        self.max_sst_entries = value;
        self
    }

    /// Sets the selected worksheet byte ceiling.
    #[must_use]
    pub const fn with_max_worksheet_scan_bytes(mut self, value: u64) -> Self {
        self.max_worksheet_scan_bytes = value;
        self
    }

    /// Sets the selected worksheet record ceiling.
    #[must_use]
    pub const fn with_max_worksheet_scan_records(mut self, value: usize) -> Self {
        self.max_worksheet_scan_records = value;
        self
    }

    /// Sets the `BoundSheet8` count ceiling.
    #[must_use]
    pub const fn with_max_sheet_count(mut self, value: usize) -> Self {
        self.max_sheet_count = value;
        self
    }

    fn validate(self) -> std::result::Result<Self, SourceBackedError> {
        if self.max_global_bytes == 0
            || self.max_materialize_bytes == 0
            || self.max_global_records == 0
            || self.max_sst_entries == 0
            || self.max_worksheet_scan_bytes == 0
            || self.max_worksheet_scan_records == 0
            || self.max_sheet_count == 0
        {
            return Err(SourceBackedError::ResourceLimit {
                resource: "source-backed XLS limit",
                observed: 0,
                maximum: 1,
            });
        }
        SharedOleFileLimits::new(self.max_input_bytes).map_err(SourceBackedError::Cfb)?;
        Ok(self)
    }
}

/// Errors specific to source-backed XLS access.
#[derive(Debug)]
pub enum SourceBackedError {
    /// An underlying positional-source error.
    Io(std::io::Error),
    /// A CFB validation or range-read error.
    Cfb(OleError),
    /// An existing XLS semantic codec rejected the source.
    Parse(Error),
    /// The source changed between two consistency fences.
    SourceChanged {
        /// Version captured before the operation.
        expected: SourceVersion,
        /// Version observed after the operation.
        observed: SourceVersion,
    },
    /// A configured or observed resource exceeded its finite ceiling.
    ResourceLimit {
        /// Resource whose ceiling was crossed.
        resource: &'static str,
        /// Observed amount.
        observed: u64,
        /// Configured ceiling.
        maximum: u64,
    },
    /// A bounded materialization allocation could not be reserved.
    Allocation {
        /// Allocation being attempted.
        resource: &'static str,
        /// Requested bytes.
        requested: u64,
    },
    /// A `FILEPASS` record was found; this read-only owner does not decrypt.
    EncryptedUnsupported,
    /// The caller's cooperative execution context cancelled the operation.
    Execution(ExecutionError),
    /// No `Workbook` or `Book` stream exists.
    WorkbookStreamMissing,
    /// The source is a valid legacy BIFF workbook whose version is not BIFF8.
    UnsupportedBiffVersion(u16),
    /// A requested worksheet does not exist.
    WorksheetNotFound(String),
    /// The workbook violates source-owner structural requirements.
    InvalidData(String),
}

impl fmt::Display for SourceBackedError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "source I/O error: {error}"),
            Self::Cfb(error) => write!(formatter, "source CFB error: {error}"),
            Self::Parse(error) => write!(formatter, "source XLS parse error: {error}"),
            Self::SourceChanged { expected, observed } => write!(
                formatter,
                "source changed from {expected:?} to {observed:?}"
            ),
            Self::ResourceLimit {
                resource,
                observed,
                maximum,
            } => write!(
                formatter,
                "source-backed XLS {resource} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::Allocation {
                resource,
                requested,
            } => write!(
                formatter,
                "source-backed XLS {resource} allocation failed for {requested} bytes"
            ),
            Self::EncryptedUnsupported => {
                formatter.write_str("encrypted XLS FILEPASS is unsupported by source-backed reads")
            },
            Self::Execution(error) => {
                write!(formatter, "source-backed XLS execution error: {error}")
            },
            Self::WorkbookStreamMissing => {
                formatter.write_str("XLS Workbook/Book stream is missing")
            },
            Self::UnsupportedBiffVersion(version) => {
                write!(
                    formatter,
                    "unsupported legacy BIFF version: 0x{version:04X}"
                )
            },
            Self::WorksheetNotFound(name) => write!(formatter, "worksheet not found: {name}"),
            Self::InvalidData(message) => formatter.write_str(message),
        }
    }
}

impl std::error::Error for SourceBackedError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(error) => Some(error),
            Self::Cfb(error) => Some(error),
            Self::Parse(error) => Some(error),
            Self::Execution(error) => Some(error),
            Self::SourceChanged { .. }
            | Self::ResourceLimit { .. }
            | Self::Allocation { .. }
            | Self::EncryptedUnsupported
            | Self::WorkbookStreamMissing
            | Self::UnsupportedBiffVersion(_)
            | Self::WorksheetNotFound(_)
            | Self::InvalidData(_) => None,
        }
    }
}

impl From<std::io::Error> for SourceBackedError {
    fn from(error: std::io::Error) -> Self {
        Self::Io(error)
    }
}

impl From<OleError> for SourceBackedError {
    fn from(error: OleError) -> Self {
        match error {
            OleError::SourceChanged { expected, observed } => {
                Self::SourceChanged { expected, observed }
            },
            OleError::LimitExceeded {
                resource,
                observed,
                maximum,
            } => Self::ResourceLimit {
                resource,
                observed,
                maximum,
            },
            other => Self::Cfb(other),
        }
    }
}

impl From<Error> for SourceBackedError {
    fn from(error: Error) -> Self {
        Self::Parse(error)
    }
}

impl From<ExecutionError> for SourceBackedError {
    fn from(error: ExecutionError) -> Self {
        Self::Execution(error)
    }
}

type Result<T> = std::result::Result<T, SourceBackedError>;

fn session_outcome<T>(result: Result<T>) -> SharedOleStreamSessionOutcome<T, SourceBackedError> {
    match result {
        Err(error @ SourceBackedError::Execution(_)) => SharedOleStreamSessionOutcome::Abort(error),
        result => SharedOleStreamSessionOutcome::Complete(result),
    }
}

#[derive(Debug, Clone)]
struct SheetEntry {
    workbook_index: usize,
    worksheet_index: Option<usize>,
    name: String,
    visibility: SheetVisibility,
    kind: SheetKind,
    start: u64,
    end: u64,
}

struct SourceInner {
    source: Arc<dyn ReadAt>,
    cfb: Arc<SharedOleFile>,
    expected_version: SourceVersion,
    workbook_path: Arc<[String]>,
    workbook_stream_len: u64,
    sheets: Box<[SheetEntry]>,
    worksheet_names: Box<[String]>,
    strings: Arc<Vec<String>>,
    formatting: Arc<Formatting>,
    encoding: Encoding,
    limits: SourceBackedLimits,
}

/// An immutable, cheaply cloned source-backed XLS workbook.
#[derive(Clone)]
pub struct SourceBackedWorkbook {
    inner: Arc<SourceInner>,
}

impl fmt::Debug for SourceBackedWorkbook {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedWorkbook")
            .field("worksheet_count", &self.inner.worksheet_names.len())
            .field("sheet_count", &self.inner.sheets.len())
            .field("workbook_stream_len", &self.inner.workbook_stream_len)
            .finish_non_exhaustive()
    }
}

/// A lifetime-free source-backed worksheet handle.
#[derive(Clone)]
pub struct SourceBackedWorksheet {
    owner: Arc<SourceInner>,
    sheet_index: usize,
}

impl fmt::Debug for SourceBackedWorksheet {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedWorksheet")
            .field("name", &self.owner.sheets[self.sheet_index].name)
            .field(
                "index",
                &self.owner.sheets[self.sheet_index]
                    .worksheet_index
                    .unwrap_or_default(),
            )
            .finish()
    }
}

/// Owned metadata for one worksheet, without source offsets or CFB IDs.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceBackedWorksheetDescriptor {
    index: usize,
    workbook_index: usize,
    name: String,
    visibility: SheetVisibility,
    kind: SheetKind,
}

impl SourceBackedWorksheetDescriptor {
    /// Zero-based worksheet index used by source-backed lookup.
    #[must_use]
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Zero-based workbook-tab index.
    #[must_use]
    pub const fn workbook_index(&self) -> usize {
        self.workbook_index
    }

    /// Worksheet name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Sheet visibility.
    #[must_use]
    pub const fn visibility(&self) -> SheetVisibility {
        self.visibility
    }

    /// Sheet kind.
    #[must_use]
    pub const fn kind(&self) -> SheetKind {
        self.kind
    }

    /// Whether the worksheet tab is visible.
    #[must_use]
    pub const fn is_visible(&self) -> bool {
        matches!(self.visibility, SheetVisibility::Visible)
    }
}

/// An owned result of one selected-cell lookup.
#[derive(Debug, Clone, PartialEq)]
pub struct SourceBackedCell {
    row: u32,
    column: u32,
    value: litchi_core::sheet::CellValue,
}

impl SourceBackedCell {
    /// Zero-based row.
    #[must_use]
    pub const fn row(&self) -> u32 {
        self.row
    }

    /// Zero-based column.
    #[must_use]
    pub const fn column(&self) -> u32 {
        self.column
    }

    /// Owned semantic cell value.
    #[must_use]
    pub const fn value(&self) -> &litchi_core::sheet::CellValue {
        &self.value
    }

    /// Consumes the lookup result and returns its value.
    #[must_use]
    pub fn into_value(self) -> litchi_core::sheet::CellValue {
        self.value
    }
}

impl SourceBackedWorkbook {
    /// Opens a positional source with default finite limits.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, SourceBackedLimits::default())
    }

    /// Opens a positional source with explicit finite limits.
    pub fn from_read_at_with_limits(
        source: Arc<dyn ReadAt>,
        limits: SourceBackedLimits,
    ) -> Result<Self> {
        limits.validate()?;
        let cfb_limits =
            SharedOleFileLimits::new(limits.max_input_bytes).map_err(SourceBackedError::Cfb)?;
        let cfb = SharedOleFile::open_with_limits(Arc::clone(&source), cfb_limits)
            .map_err(SourceBackedError::from)?;
        Self::from_shared_ole_file_with_limits(Arc::new(cfb), limits)
    }

    /// Opens a filesystem path through the positional `FileSource` adapter.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        let source = Arc::new(FileSource::open(path).map_err(SourceBackedError::Io)?);
        Self::from_read_at(source)
    }

    /// Opens a filesystem path with explicit finite limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(
        path: impl AsRef<Path>,
        limits: SourceBackedLimits,
    ) -> Result<Self> {
        let source = Arc::new(FileSource::open(path).map_err(SourceBackedError::Io)?);
        Self::from_read_at_with_limits(source, limits)
    }

    /// Advanced compatibility adapter to the existing eager semantic owner.
    ///
    /// This fallback reads the original positional source once, rather than
    /// reopening the path or exposing raw CFB bytes. The copy is bounded by
    /// [`SourceBackedLimits::max_materialize_bytes`]. It is not ordinary
    /// source-backed CRUD; callers should prefer the source-backed selectors
    /// when they do not specifically need the eager `Workbook` API.
    pub fn materialize_eager(&self) -> Result<Workbook<Cursor<Vec<u8>>>> {
        self.materialize_eager_impl(None)
    }

    /// Advanced compatibility adapter with cooperative cancellation checks.
    ///
    /// Cancellation is checked before the copy, between bounded source reads,
    /// and before and after eager semantic parsing. The existing eager parser
    /// is not execution-aware, so cancellation cannot interrupt it mid-parse;
    /// its in-memory input is bounded by `max_materialize_bytes` and a final
    /// cancellation check runs after parsing.
    pub fn materialize_eager_with_execution(
        &self,
        execution: &ExecutionContext,
    ) -> Result<Workbook<Cursor<Vec<u8>>>> {
        self.materialize_eager_impl(Some(execution))
    }

    fn materialize_eager_impl(
        &self,
        execution: Option<&ExecutionContext>,
    ) -> Result<Workbook<Cursor<Vec<u8>>>> {
        if let Some(execution) = execution {
            execution.check()?;
        }
        self.inner.ensure_current()?;
        let source_len = self.inner.source.len()?;
        self.inner.ensure_current()?;
        if source_len > self.inner.limits.max_materialize_bytes {
            return Err(SourceBackedError::ResourceLimit {
                resource: "materialization bytes",
                observed: source_len,
                maximum: self.inner.limits.max_materialize_bytes,
            });
        }
        let length =
            usize::try_from(source_len).map_err(|_error| SourceBackedError::ResourceLimit {
                resource: "materialization address space",
                observed: source_len,
                maximum: u64::try_from(usize::MAX).unwrap_or(u64::MAX),
            })?;
        let mut bytes = allocate_materialization_buffer(length, source_len)?;

        let mut offset = 0_usize;
        while offset < length {
            if let Some(execution) = execution {
                execution.check()?;
            }
            let count = (length - offset).min(MATERIALIZE_CHUNK_BYTES);
            let mut filled = 0_usize;
            while filled < count {
                if let Some(execution) = execution {
                    execution.check()?;
                }
                let read = match self.inner.source.read_at(
                    u64::try_from(offset + filled).map_err(|_error| {
                        SourceBackedError::InvalidData(
                            "materialization offset does not fit u64".into(),
                        )
                    })?,
                    &mut bytes[offset + filled..offset + count],
                ) {
                    Ok(read) => read,
                    Err(error) => {
                        self.inner.ensure_current()?;
                        return Err(SourceBackedError::Io(error));
                    },
                };
                if read == 0 {
                    self.inner.ensure_current()?;
                    return Err(SourceBackedError::InvalidData(
                        "source ended during eager materialization".into(),
                    ));
                }
                let remaining = count - filled;
                if read > remaining {
                    self.inner.ensure_current()?;
                    return Err(SourceBackedError::InvalidData(
                        "source returned more bytes than requested".into(),
                    ));
                }
                filled += read;
                self.inner.ensure_current()?;
            }
            offset += count;
        }

        if let Some(execution) = execution {
            execution.check()?;
        }
        self.inner.ensure_current()?;
        let workbook = match Workbook::new(Cursor::new(bytes)) {
            Ok(workbook) => workbook,
            Err(error) => {
                self.inner.ensure_current()?;
                return Err(SourceBackedError::Parse(error));
            },
        };
        if let Some(execution) = execution {
            execution.check()?;
        }
        self.inner.ensure_current()?;
        Ok(workbook)
    }

    /// Adopts an already-indexed positional CFB file.
    pub(crate) fn from_shared_ole_file_with_limits(
        cfb: Arc<SharedOleFile>,
        limits: SourceBackedLimits,
    ) -> Result<Self> {
        limits.validate()?;
        let source = cfb.source_arc();
        let expected_version = cfb.source_version().map_err(SourceBackedError::from)?;
        ensure_current_parts(&source, &cfb, expected_version)?;
        let file_size = cfb.file_size();
        if file_size > limits.max_input_bytes {
            return Err(SourceBackedError::ResourceLimit {
                resource: "input bytes",
                observed: file_size,
                maximum: limits.max_input_bytes,
            });
        }
        let (workbook_path, workbook_stream_len) =
            select_workbook_stream(&cfb, &source, expected_version)?;
        let mut parsed = parse_globals(&cfb, &workbook_path, workbook_stream_len, limits)?;
        validate_sheet_offsets(workbook_stream_len, parsed.global_end, &mut parsed.sheets)?;
        let mut worksheet_names = Vec::new();
        for sheet in &parsed.sheets {
            if sheet.kind == SheetKind::WorksheetOrDialog {
                worksheet_names.try_reserve(1).map_err(|_error| {
                    SourceBackedError::InvalidData("worksheet name allocation failed".into())
                })?;
                worksheet_names.push(sheet.name.clone());
            }
        }
        let inner = Arc::new(SourceInner {
            source,
            cfb,
            expected_version,
            workbook_path: workbook_path.into_boxed_slice().into(),
            workbook_stream_len,
            sheets: parsed.sheets.into_boxed_slice(),
            worksheet_names: worksheet_names.into_boxed_slice(),
            strings: Arc::new(parsed.strings),
            formatting: Arc::new(parsed.formatting),
            encoding: parsed.encoding,
            limits,
        });
        inner.ensure_current()?;
        Ok(Self { inner })
    }

    /// Number of all workbook tabs, including chart and macro sheets.
    pub fn sheet_count(&self) -> Result<usize> {
        self.metadata(|inner| inner.sheets.len())
    }

    /// Number of ordinary worksheet tabs.
    pub fn worksheet_count(&self) -> Result<usize> {
        self.metadata(|inner| inner.worksheet_names.len())
    }

    /// Names of ordinary worksheet tabs in worksheet-index order.
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        self.metadata(|inner| inner.worksheet_names.to_vec())
    }

    /// Metadata descriptors for ordinary worksheet tabs.
    pub fn worksheet_descriptors(&self) -> Result<Vec<SourceBackedWorksheetDescriptor>> {
        self.metadata(|inner| descriptors(&inner.sheets))
    }

    /// Alias returning the ordinary worksheet descriptors in tab order.
    pub fn sheets(&self) -> Result<Vec<SourceBackedWorksheetDescriptor>> {
        self.worksheet_descriptors()
    }

    /// Returns one ordinary worksheet descriptor by zero-based index.
    pub fn worksheet_descriptor(
        &self,
        index: usize,
    ) -> Result<Option<SourceBackedWorksheetDescriptor>> {
        self.metadata(|inner| {
            inner
                .sheets
                .iter()
                .find(|sheet| sheet.worksheet_index == Some(index))
                .map(|sheet| SourceBackedWorksheetDescriptor {
                    index,
                    workbook_index: sheet.workbook_index,
                    name: sheet.name.clone(),
                    visibility: sheet.visibility,
                    kind: sheet.kind,
                })
        })
    }

    /// Returns one ordinary worksheet by zero-based worksheet index.
    pub fn worksheet_by_index(&self, index: usize) -> Result<Option<SourceBackedWorksheet>> {
        self.metadata(|inner| {
            let sheet_index = inner
                .sheets
                .iter()
                .position(|sheet| sheet.worksheet_index == Some(index));
            sheet_index.map(|sheet_index| SourceBackedWorksheet {
                owner: Arc::clone(&self.inner),
                sheet_index,
            })
        })
    }

    /// Alias for [`Self::worksheet_by_index`].
    pub fn worksheet(&self, index: usize) -> Result<Option<SourceBackedWorksheet>> {
        self.worksheet_by_index(index)
    }

    /// Returns one ordinary worksheet by case-insensitive name.
    pub fn worksheet_by_name(&self, name: &str) -> Result<Option<SourceBackedWorksheet>> {
        self.metadata(|inner| {
            let sheet_index = inner.sheets.iter().position(|sheet| {
                sheet.worksheet_index.is_some() && sheet.name.eq_ignore_ascii_case(name)
            });
            sheet_index.map(|sheet_index| SourceBackedWorksheet {
                owner: Arc::clone(&self.inner),
                sheet_index,
            })
        })
    }

    /// Iterates ordinary worksheet handles without reading worksheet payloads.
    pub fn worksheets(&self) -> Result<Vec<SourceBackedWorksheet>> {
        self.metadata(|inner| {
            inner
                .sheets
                .iter()
                .enumerate()
                .filter_map(|(sheet_index, sheet)| {
                    sheet.worksheet_index.map(|_| SourceBackedWorksheet {
                        owner: Arc::clone(&self.inner),
                        sheet_index,
                    })
                })
                .collect()
        })
    }

    /// Looks up one cell by worksheet index and zero-based coordinates.
    pub fn cell_by_index(
        &self,
        worksheet_index: usize,
        row: u32,
        column: u32,
    ) -> Result<Option<SourceBackedCell>> {
        let worksheet = self
            .worksheet_by_index(worksheet_index)?
            .ok_or_else(|| SourceBackedError::WorksheetNotFound(worksheet_index.to_string()))?;
        worksheet.cell(row, column)
    }

    /// Looks up one cell value by worksheet index and zero-based coordinates.
    pub fn cell_value_by_index(
        &self,
        worksheet_index: usize,
        row: u32,
        column: u32,
    ) -> Result<Option<litchi_core::sheet::CellValue>> {
        Ok(self
            .cell_by_index(worksheet_index, row, column)?
            .map(SourceBackedCell::into_value))
    }

    /// Looks up one cell by worksheet index with cooperative cancellation.
    pub fn cell_by_index_with_execution(
        &self,
        worksheet_index: usize,
        row: u32,
        column: u32,
        execution: &ExecutionContext,
    ) -> Result<Option<SourceBackedCell>> {
        execution.check().map_err(SourceBackedError::from)?;
        let worksheet = self
            .worksheet_by_index(worksheet_index)?
            .ok_or_else(|| SourceBackedError::WorksheetNotFound(worksheet_index.to_string()))?;
        worksheet.cell_with_execution(row, column, execution)
    }

    /// Looks up one cell value by worksheet index with cooperative
    /// cancellation.
    pub fn cell_value_by_index_with_execution(
        &self,
        worksheet_index: usize,
        row: u32,
        column: u32,
        execution: &ExecutionContext,
    ) -> Result<Option<litchi_core::sheet::CellValue>> {
        Ok(self
            .cell_by_index_with_execution(worksheet_index, row, column, execution)?
            .map(SourceBackedCell::into_value))
    }

    /// Alias for [`Self::cell_value_by_index`].
    pub fn cell_value(
        &self,
        worksheet_index: usize,
        row: u32,
        column: u32,
    ) -> Result<Option<litchi_core::sheet::CellValue>> {
        self.cell_value_by_index(worksheet_index, row, column)
    }

    /// Alias for [`Self::cell_value_by_index_with_execution`].
    pub fn cell_value_with_execution(
        &self,
        worksheet_index: usize,
        row: u32,
        column: u32,
        execution: &ExecutionContext,
    ) -> Result<Option<litchi_core::sheet::CellValue>> {
        self.cell_value_by_index_with_execution(worksheet_index, row, column, execution)
    }

    /// Looks up one cell by case-insensitive worksheet name and coordinates.
    pub fn cell_by_name(
        &self,
        name: &str,
        row: u32,
        column: u32,
    ) -> Result<Option<SourceBackedCell>> {
        let worksheet = self
            .worksheet_by_name(name)?
            .ok_or_else(|| SourceBackedError::WorksheetNotFound(name.to_string()))?;
        worksheet.cell(row, column)
    }

    /// Looks up one cell value by case-insensitive worksheet name.
    pub fn cell_value_by_name(
        &self,
        name: &str,
        row: u32,
        column: u32,
    ) -> Result<Option<litchi_core::sheet::CellValue>> {
        Ok(self
            .cell_by_name(name, row, column)?
            .map(SourceBackedCell::into_value))
    }

    /// Looks up one cell by worksheet name with cooperative cancellation.
    pub fn cell_by_name_with_execution(
        &self,
        name: &str,
        row: u32,
        column: u32,
        execution: &ExecutionContext,
    ) -> Result<Option<SourceBackedCell>> {
        execution.check().map_err(SourceBackedError::from)?;
        let worksheet = self
            .worksheet_by_name(name)?
            .ok_or_else(|| SourceBackedError::WorksheetNotFound(name.to_string()))?;
        worksheet.cell_with_execution(row, column, execution)
    }

    /// Looks up one cell value by worksheet name with cooperative
    /// cancellation.
    pub fn cell_value_by_name_with_execution(
        &self,
        name: &str,
        row: u32,
        column: u32,
        execution: &ExecutionContext,
    ) -> Result<Option<litchi_core::sheet::CellValue>> {
        Ok(self
            .cell_by_name_with_execution(name, row, column, execution)?
            .map(SourceBackedCell::into_value))
    }

    /// Returns the parsed workbook date system.
    pub fn date_system(&self) -> Result<DateSystem> {
        self.metadata(|inner| inner.formatting.date_system())
    }

    /// Returns the source version captured at open after a consistency check.
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.ensure_current()?;
        Ok(self.inner.expected_version)
    }

    fn ensure_current(&self) -> Result<()> {
        self.inner.ensure_current()
    }

    fn metadata<T>(&self, operation: impl FnOnce(&SourceInner) -> T) -> Result<T> {
        self.ensure_current()?;
        let value = operation(&self.inner);
        self.ensure_current()?;
        Ok(value)
    }
}

fn allocate_materialization_buffer(length: usize, requested: u64) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(length)
        .map_err(|_error| SourceBackedError::Allocation {
            resource: "eager materialization buffer",
            requested,
        })?;
    bytes.resize(length, 0);
    Ok(bytes)
}

#[cfg(test)]
mod materialization_tests {
    use super::*;

    #[test]
    fn reserve_failure_is_reported_as_typed_allocation_error() {
        let error = allocate_materialization_buffer(usize::MAX, u64::MAX).unwrap_err();
        assert!(matches!(
            error,
            SourceBackedError::Allocation {
                resource: "eager materialization buffer",
                requested: u64::MAX,
            }
        ));
    }
}

impl SourceInner {
    fn ensure_current(&self) -> Result<()> {
        ensure_current_parts(&self.source, &self.cfb, self.expected_version)
    }
}

fn ensure_current_parts(
    source: &Arc<dyn ReadAt>,
    cfb: &SharedOleFile,
    expected_version: SourceVersion,
) -> Result<()> {
    let observed = source.version().map_err(SourceBackedError::Io)?;
    if observed != expected_version {
        return Err(SourceBackedError::SourceChanged {
            expected: expected_version,
            observed,
        });
    }
    cfb.source_version().map_err(SourceBackedError::from)?;
    let observed = source.version().map_err(SourceBackedError::Io)?;
    if observed != expected_version {
        return Err(SourceBackedError::SourceChanged {
            expected: expected_version,
            observed,
        });
    }
    Ok(())
}

impl SourceBackedWorksheet {
    /// Zero-based worksheet index.
    pub fn index(&self) -> Result<usize> {
        self.metadata(|sheet| sheet.worksheet_index.unwrap_or_default())
    }

    /// Worksheet name.
    pub fn name(&self) -> Result<String> {
        self.metadata(|sheet| sheet.name.clone())
    }

    /// Workbook-tab index.
    pub fn workbook_index(&self) -> Result<usize> {
        self.metadata(|sheet| sheet.workbook_index)
    }

    /// Sheet visibility.
    pub fn visibility(&self) -> Result<SheetVisibility> {
        self.metadata(|sheet| sheet.visibility)
    }

    /// Sheet kind.
    pub fn kind(&self) -> Result<SheetKind> {
        self.metadata(|sheet| sheet.kind)
    }

    /// Worksheet metadata descriptor without source offsets.
    pub fn descriptor(&self) -> Result<SourceBackedWorksheetDescriptor> {
        self.metadata(descriptor)
    }

    /// Reads one cell by zero-based row and column.
    pub fn cell(&self, row: u32, column: u32) -> Result<Option<SourceBackedCell>> {
        query_cell(&self.owner, self.sheet_index, row, column, None)
    }

    /// Reads one cell by zero-based row and column with cooperative
    /// cancellation.
    pub fn cell_with_execution(
        &self,
        row: u32,
        column: u32,
        execution: &ExecutionContext,
    ) -> Result<Option<SourceBackedCell>> {
        query_cell(&self.owner, self.sheet_index, row, column, Some(execution))
    }

    /// Reads one cell value by zero-based row and column.
    pub fn cell_value(
        &self,
        row: u32,
        column: u32,
    ) -> Result<Option<litchi_core::sheet::CellValue>> {
        Ok(self.cell(row, column)?.map(SourceBackedCell::into_value))
    }

    /// Reads one cell value by zero-based row and column with cooperative
    /// cancellation.
    pub fn cell_value_with_execution(
        &self,
        row: u32,
        column: u32,
        execution: &ExecutionContext,
    ) -> Result<Option<litchi_core::sheet::CellValue>> {
        Ok(self
            .cell_with_execution(row, column, execution)?
            .map(SourceBackedCell::into_value))
    }

    fn metadata<T>(&self, operation: impl FnOnce(&SheetEntry) -> T) -> Result<T> {
        self.owner.ensure_current()?;
        let sheet =
            self.owner.sheets.get(self.sheet_index).ok_or_else(|| {
                SourceBackedError::WorksheetNotFound(self.sheet_index.to_string())
            })?;
        let value = operation(sheet);
        self.owner.ensure_current()?;
        Ok(value)
    }
}

struct ParsedGlobals {
    global_end: u64,
    sheets: Vec<SheetEntry>,
    strings: Vec<String>,
    formatting: Formatting,
    encoding: Encoding,
}

fn descriptor(sheet: &SheetEntry) -> SourceBackedWorksheetDescriptor {
    SourceBackedWorksheetDescriptor {
        index: sheet.worksheet_index.unwrap_or_default(),
        workbook_index: sheet.workbook_index,
        name: sheet.name.clone(),
        visibility: sheet.visibility,
        kind: sheet.kind,
    }
}

fn descriptors(sheets: &[SheetEntry]) -> Vec<SourceBackedWorksheetDescriptor> {
    sheets
        .iter()
        .filter(|sheet| sheet.worksheet_index.is_some())
        .map(descriptor)
        .collect()
}

fn select_workbook_stream(
    cfb: &SharedOleFile,
    source: &Arc<dyn ReadAt>,
    expected_version: SourceVersion,
) -> Result<(Vec<String>, u64)> {
    ensure_current_parts(source, cfb, expected_version)?;
    for name in ["Workbook", "Book"] {
        let path = vec![name.to_string()];
        let refs = [name];
        match cfb.stream_len(&refs) {
            Ok(length) => {
                ensure_current_parts(source, cfb, expected_version)?;
                return Ok((path, length));
            },
            Err(OleError::StreamNotFound) => {
                ensure_current_parts(source, cfb, expected_version)?;
            },
            Err(error) => {
                ensure_current_parts(source, cfb, expected_version)?;
                return Err(SourceBackedError::from(error));
            },
        }
    }
    ensure_current_parts(source, cfb, expected_version)?;
    Err(SourceBackedError::WorkbookStreamMissing)
}

fn parse_globals(
    cfb: &SharedOleFile,
    path: &[String],
    stream_len: u64,
    limits: SourceBackedLimits,
) -> Result<ParsedGlobals> {
    let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
    let (offset, record_count) = cfb.with_stream_session_at(&refs, 0, |session| {
        session_outcome(preflight_global_headers(session, stream_len, limits))
    })?;

    let global_len =
        usize::try_from(offset).map_err(|_error| SourceBackedError::ResourceLimit {
            resource: "global address space",
            observed: offset,
            maximum: usize::MAX as u64,
        })?;
    let bytes = read_range(cfb, &refs, 0, global_len)?;

    let biff_limits = BiffLimits {
        max_records: limits.max_global_records,
        max_record_bytes: litchi_biff::MAX_RECORD_BYTES,
        max_input_bytes: usize::try_from(limits.max_global_bytes).unwrap_or(usize::MAX),
        max_output_bytes: usize::MAX,
    };
    let mut records = Vec::new();
    records
        .try_reserve(record_count)
        .map_err(|_error| SourceBackedError::Allocation {
            resource: "global records",
            requested: record_count as u64,
        })?;
    for record in BiffRecords::with_limits(&bytes, biff_limits).map_err(map_biff_error)? {
        records.push(record.map_err(map_biff_error)?);
    }
    let first = records
        .first()
        .ok_or_else(|| SourceBackedError::InvalidData("Workbook globals are empty".into()))?;
    if first.kind().get() != BOF {
        return Err(SourceBackedError::InvalidData(
            "Workbook globals do not start with BOF".into(),
        ));
    }
    let bof = BofRecord::parse(first.payload()).map_err(SourceBackedError::Parse)?;
    if bof.version as u16 != BIFF8 {
        return Err(SourceBackedError::UnsupportedBiffVersion(
            bof.version as u16,
        ));
    }
    if first.payload().len() < 4 {
        return Err(SourceBackedError::Parse(Error::InvalidLength {
            expected: 4,
            found: first.payload().len(),
        }));
    }
    let substream_type = u16::from_le_bytes([first.payload()[2], first.payload()[3]]);
    if substream_type != WORKBOOK_BOF_TYPE {
        return Err(SourceBackedError::InvalidData(
            "Workbook globals BOF has an invalid substream type".into(),
        ));
    }
    let Some(last) = records.last() else {
        return Err(SourceBackedError::InvalidData(
            "Workbook globals do not end with EOF".into(),
        ));
    };
    if last.kind().get() != EOF {
        return Err(SourceBackedError::InvalidData(
            "Workbook globals do not end with EOF".into(),
        ));
    }
    if !last.payload().is_empty() {
        return Err(SourceBackedError::InvalidData(
            "Workbook globals EOF has a non-empty payload".into(),
        ));
    }

    let mut encoding =
        Encoding::from_codepage(DEFAULT_CODEPAGE).map_err(SourceBackedError::Parse)?;
    // BoundSheet8 names use the workbook CODEPAGE.  Keep the raw payloads
    // until all globals have been framed so a late CODEPAGE has the same
    // semantics as the eager parser.
    let mut bound_payloads = Vec::<&[u8]>::new();
    let mut sst_refs = Vec::<RecordRef<'_>>::new();
    let mut sst_seen = false;
    let mut i = 0_usize;
    while i < records.len() {
        let record = records[i];
        match record.kind().get() {
            FILEPASS => return Err(SourceBackedError::EncryptedUnsupported),
            CODEPAGE => {
                if record.payload().len() != 2 {
                    return Err(SourceBackedError::Parse(Error::InvalidLength {
                        expected: 2,
                        found: record.payload().len(),
                    }));
                }
                let codepage = u16::from_le_bytes([record.payload()[0], record.payload()[1]]);
                encoding = Encoding::from_codepage(codepage).map_err(SourceBackedError::Parse)?;
            },
            BOUND_SHEET => {
                if bound_payloads.len() >= limits.max_sheet_count {
                    return Err(SourceBackedError::ResourceLimit {
                        resource: "sheet count",
                        observed: (bound_payloads.len() + 1) as u64,
                        maximum: limits.max_sheet_count as u64,
                    });
                }
                bound_payloads
                    .try_reserve(1)
                    .map_err(|_error| SourceBackedError::Allocation {
                        resource: "BoundSheet8 payloads",
                        requested: 1,
                    })?;
                bound_payloads.push(record.payload());
            },
            SST => {
                if sst_seen {
                    return Err(SourceBackedError::InvalidData(
                        "Workbook globals contain multiple SST records".into(),
                    ));
                }
                sst_seen = true;
                sst_refs.push(record);
                while records
                    .get(i + 1)
                    .is_some_and(|next| next.kind().get() == CONTINUE)
                {
                    i += 1;
                    sst_refs.push(records[i]);
                }
                if record.payload().len() < 8 {
                    return Err(SourceBackedError::Parse(Error::InvalidLength {
                        expected: 8,
                        found: record.payload().len(),
                    }));
                }
                let total = u32::from_le_bytes([
                    record.payload()[0],
                    record.payload()[1],
                    record.payload()[2],
                    record.payload()[3],
                ]);
                let unique = u32::from_le_bytes([
                    record.payload()[4],
                    record.payload()[5],
                    record.payload()[6],
                    record.payload()[7],
                ]);
                let maximum = limits.max_sst_entries as u64;
                if u64::from(total) > maximum || u64::from(unique) > maximum {
                    return Err(SourceBackedError::ResourceLimit {
                        resource: "SST entries",
                        observed: u64::from(total.max(unique)),
                        maximum,
                    });
                }
            },
            _ => {},
        }
        i += 1;
    }
    if bound_payloads.is_empty() {
        return Err(SourceBackedError::InvalidData(
            "Workbook globals contain no BoundSheet8 records".into(),
        ));
    }
    let mut bounds = Vec::new();
    bounds
        .try_reserve(bound_payloads.len())
        .map_err(|_error| SourceBackedError::Allocation {
            resource: "BoundSheet8 records",
            requested: bound_payloads.len() as u64,
        })?;
    for payload in bound_payloads {
        bounds.push(BoundSheetRecord::parse(payload, &encoding).map_err(SourceBackedError::Parse)?);
    }
    let mut names = HashSet::new();
    names
        .try_reserve(bounds.len())
        .map_err(|_error| SourceBackedError::Allocation {
            resource: "BoundSheet8 names",
            requested: bounds.len() as u64,
        })?;
    for bound in &bounds {
        if !names.insert(bound.name.to_lowercase()) {
            return Err(SourceBackedError::Parse(Error::InvalidRecord {
                record_type: BOUND_SHEET,
                message: format!(
                    "duplicate case-insensitive BoundSheet8 name: {:?}",
                    bound.name
                ),
            }));
        }
    }
    let mut tolerance = ToleranceLog::new(Leniency::Strict);
    let formatting =
        Formatting::parse_globals(&records, &mut tolerance).map_err(SourceBackedError::Parse)?;
    let strings = if sst_refs.is_empty() {
        Vec::new()
    } else {
        crate::records::SharedStringTable::parse_from_records(&sst_refs, &encoding)
            .map_err(SourceBackedError::Parse)?
            .strings
    };

    let mut sheets = Vec::new();
    sheets
        .try_reserve(bounds.len())
        .map_err(|_error| SourceBackedError::Allocation {
            resource: "source-backed sheet descriptors",
            requested: bounds.len() as u64,
        })?;
    let mut worksheet_index = 0_usize;
    for (workbook_index, sheet) in bounds.into_iter().enumerate() {
        let (visibility, kind) = (
            match sheet.visible {
                crate::records::SheetVisible::Visible => SheetVisibility::Visible,
                crate::records::SheetVisible::Hidden => SheetVisibility::Hidden,
                crate::records::SheetVisible::VeryHidden => SheetVisibility::VeryHidden,
            },
            match sheet.sheet_type {
                SheetType::WorkSheet => SheetKind::WorksheetOrDialog,
                SheetType::MacroSheet => SheetKind::MacroSheet,
                SheetType::ChartSheet => SheetKind::ChartSheet,
                SheetType::VBModule => SheetKind::VbaModule,
            },
        );
        let current_worksheet_index = (kind == SheetKind::WorksheetOrDialog).then(|| {
            let value = worksheet_index;
            worksheet_index += 1;
            value
        });
        sheets.push(SheetEntry {
            workbook_index,
            worksheet_index: current_worksheet_index,
            name: sheet.name,
            visibility,
            kind,
            start: u64::from(sheet.position),
            end: stream_len,
        });
    }
    Ok(ParsedGlobals {
        global_end: offset,
        sheets,
        strings,
        formatting,
        encoding,
    })
}

fn preflight_global_headers(
    session: &mut SharedOleStreamSession<'_>,
    stream_len: u64,
    limits: SourceBackedLimits,
) -> Result<(u64, usize)> {
    let mut offset = 0_u64;
    let mut record_count = 0_usize;
    loop {
        let header_end = offset.checked_add(4).ok_or_else(|| {
            SourceBackedError::InvalidData("BIFF global header offset overflows".into())
        })?;
        if header_end > stream_len {
            return Err(SourceBackedError::InvalidData(
                "truncated BIFF global record header".into(),
            ));
        }
        let mut header = [0_u8; 4];
        session
            .read_exact(&mut header)
            .map_err(SourceBackedError::from)?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        if kind == FILEPASS {
            return Err(SourceBackedError::EncryptedUnsupported);
        }
        let payload_len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        if payload_len > litchi_biff::MAX_RECORD_BYTES {
            return Err(SourceBackedError::ResourceLimit {
                resource: "BIFF record bytes",
                observed: payload_len as u64,
                maximum: litchi_biff::MAX_RECORD_BYTES as u64,
            });
        }
        let frame_len = 4_u64
            .checked_add(payload_len as u64)
            .ok_or_else(|| SourceBackedError::InvalidData("BIFF global frame overflows".into()))?;
        let end = offset
            .checked_add(frame_len)
            .ok_or_else(|| SourceBackedError::InvalidData("BIFF global offset overflows".into()))?;
        if end > stream_len {
            return Err(SourceBackedError::InvalidData(
                "BIFF global record exceeds Workbook stream".into(),
            ));
        }
        record_count = record_count.checked_add(1).ok_or({
            SourceBackedError::ResourceLimit {
                resource: "global records",
                observed: u64::MAX,
                maximum: limits.max_global_records as u64,
            }
        })?;
        if record_count > limits.max_global_records {
            return Err(SourceBackedError::ResourceLimit {
                resource: "global records",
                observed: record_count as u64,
                maximum: limits.max_global_records as u64,
            });
        }
        if end > limits.max_global_bytes {
            return Err(SourceBackedError::ResourceLimit {
                resource: "global bytes",
                observed: end,
                maximum: limits.max_global_bytes,
            });
        }
        session
            .skip_forward(payload_len as u64)
            .map_err(SourceBackedError::from)?;
        offset = end;
        if kind == EOF {
            if payload_len != 0 {
                return Err(SourceBackedError::InvalidData(
                    "Workbook globals EOF has a non-empty payload".into(),
                ));
            }
            break;
        }
    }
    Ok((offset, record_count))
}

fn validate_sheet_offsets(
    stream_len: u64,
    global_end: u64,
    sheets: &mut [SheetEntry],
) -> Result<()> {
    let mut order = (0..sheets.len()).collect::<Vec<_>>();
    order.sort_unstable_by_key(|index| sheets[*index].start);
    for pair in order.windows(2) {
        if sheets[pair[0]].start == sheets[pair[1]].start {
            return Err(SourceBackedError::InvalidData(
                "duplicate BoundSheet8 stream offsets".into(),
            ));
        }
    }
    for (position, index) in order.iter().copied().enumerate() {
        let start = sheets[index].start;
        let upper_bound = order
            .get(position + 1)
            .map_or(stream_len, |next| sheets[*next].start);
        if start < global_end
            || start >= upper_bound
            || start.checked_add(4).is_none_or(|value| value > upper_bound)
            || upper_bound > stream_len
        {
            return Err(SourceBackedError::InvalidData(
                "BoundSheet8 stream offset is outside the Workbook stream".into(),
            ));
        }
        sheets[index].end = upper_bound;
    }
    Ok(())
}

fn read_range(cfb: &SharedOleFile, path: &[&str], offset: u64, length: usize) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|_error| SourceBackedError::Allocation {
            resource: "source-backed range buffer",
            requested: length as u64,
        })?;
    output.resize(length, 0);
    cfb.read_stream_range(path, offset, &mut output)
        .map_err(SourceBackedError::from)?;
    Ok(output)
}

struct WorksheetFrame {
    kind: u16,
    payload_len: usize,
}

struct WorksheetScan<'a> {
    position: u64,
    upper_bound: u64,
    scanned_bytes: u64,
    scanned_records: usize,
    limits: SourceBackedLimits,
    execution: Option<&'a ExecutionContext>,
}

impl<'a> WorksheetScan<'a> {
    fn new(
        start: u64,
        upper_bound: u64,
        limits: SourceBackedLimits,
        execution: Option<&'a ExecutionContext>,
    ) -> Self {
        Self {
            position: start,
            upper_bound,
            scanned_bytes: 0,
            scanned_records: 0,
            limits,
            execution,
        }
    }

    fn next_frame(&mut self, session: &mut SharedOleStreamSession<'_>) -> Result<WorksheetFrame> {
        self.check_execution()?;
        let cursor_position = self.position;
        let header_end = cursor_position.checked_add(4).ok_or_else(|| {
            SourceBackedError::InvalidData("BIFF worksheet offset overflows".into())
        })?;
        if header_end > self.upper_bound {
            return Err(SourceBackedError::InvalidData(
                "BIFF worksheet has no complete record header before its boundary".into(),
            ));
        }
        let mut header = [0_u8; 4];
        session
            .read_exact(&mut header)
            .map_err(SourceBackedError::from)?;
        self.position = header_end;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let payload_len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        if payload_len > litchi_biff::MAX_RECORD_BYTES {
            return Err(SourceBackedError::ResourceLimit {
                resource: "BIFF record bytes",
                observed: payload_len as u64,
                maximum: litchi_biff::MAX_RECORD_BYTES as u64,
            });
        }
        let frame_len = 4_u64.checked_add(payload_len as u64).ok_or_else(|| {
            SourceBackedError::InvalidData("BIFF worksheet frame overflows".into())
        })?;
        let next = cursor_position.checked_add(frame_len).ok_or_else(|| {
            SourceBackedError::InvalidData("BIFF worksheet offset overflows".into())
        })?;
        if next > self.upper_bound {
            return Err(SourceBackedError::InvalidData(
                "BIFF worksheet record exceeds its BoundSheet boundary".into(),
            ));
        }
        let records = self.scanned_records.checked_add(1).ok_or({
            SourceBackedError::ResourceLimit {
                resource: "worksheet scan records",
                observed: u64::MAX,
                maximum: self.limits.max_worksheet_scan_records as u64,
            }
        })?;
        if records > self.limits.max_worksheet_scan_records {
            return Err(SourceBackedError::ResourceLimit {
                resource: "worksheet scan records",
                observed: records as u64,
                maximum: self.limits.max_worksheet_scan_records as u64,
            });
        }
        let bytes = self.scanned_bytes.checked_add(frame_len).ok_or({
            SourceBackedError::ResourceLimit {
                resource: "worksheet scan bytes",
                observed: u64::MAX,
                maximum: self.limits.max_worksheet_scan_bytes,
            }
        })?;
        if bytes > self.limits.max_worksheet_scan_bytes {
            return Err(SourceBackedError::ResourceLimit {
                resource: "worksheet scan bytes",
                observed: bytes,
                maximum: self.limits.max_worksheet_scan_bytes,
            });
        }
        self.scanned_records = records;
        self.scanned_bytes = bytes;
        Ok(WorksheetFrame { kind, payload_len })
    }

    fn read_payload(
        &mut self,
        session: &mut SharedOleStreamSession<'_>,
        frame: &WorksheetFrame,
    ) -> Result<Vec<u8>> {
        self.check_execution()?;
        let mut payload = Vec::new();
        payload
            .try_reserve_exact(frame.payload_len)
            .map_err(|_error| SourceBackedError::Allocation {
                resource: "source-backed worksheet payload",
                requested: frame.payload_len as u64,
            })?;
        payload.resize(frame.payload_len, 0);
        session
            .read_exact(&mut payload)
            .map_err(SourceBackedError::from)?;
        self.position = self
            .position
            .checked_add(frame.payload_len as u64)
            .ok_or_else(|| {
                SourceBackedError::InvalidData("BIFF worksheet offset overflows".into())
            })?;
        Ok(payload)
    }

    fn skip_payload(
        &mut self,
        session: &mut SharedOleStreamSession<'_>,
        frame: &WorksheetFrame,
    ) -> Result<()> {
        self.check_execution()?;
        session
            .skip_forward(frame.payload_len as u64)
            .map_err(SourceBackedError::from)?;
        self.position = self
            .position
            .checked_add(frame.payload_len as u64)
            .ok_or_else(|| {
                SourceBackedError::InvalidData("BIFF worksheet offset overflows".into())
            })?;
        Ok(())
    }

    fn check_execution(&self) -> Result<()> {
        self.execution.map_or(Ok(()), |context| {
            context.check().map_err(SourceBackedError::from)
        })
    }
}

/*
 * The methods below intentionally remain in the source-backed owner rather
 * than retaining a session: session lifetimes are scoped to one query.
 */
fn validate_worksheet_bof(payload: &[u8]) -> Result<()> {
    let bof = BofRecord::parse(payload).map_err(SourceBackedError::Parse)?;
    if bof.version as u16 != BIFF8 {
        return Err(SourceBackedError::Parse(Error::UnsupportedBiffVersion(
            bof.version as u16,
        )));
    }
    if payload.len() < 4 {
        return Err(SourceBackedError::Parse(Error::InvalidLength {
            expected: 4,
            found: payload.len(),
        }));
    }
    let substream_type = u16::from_le_bytes([payload[2], payload[3]]);
    if substream_type != WORKSHEET_BOF_TYPE {
        return Err(SourceBackedError::InvalidData(
            "BoundSheet8 position does not point to a worksheet BOF".into(),
        ));
    }
    Ok(())
}

fn query_cell(
    owner: &Arc<SourceInner>,
    sheet_index: usize,
    row: u32,
    column: u32,
    execution: Option<&ExecutionContext>,
) -> Result<Option<SourceBackedCell>> {
    if let Some(context) = execution {
        context.check().map_err(SourceBackedError::from)?;
    }
    owner.ensure_current()?;
    if row > u32::from(u16::MAX) || column > u32::from(u8::MAX) {
        owner.ensure_current()?;
        return Ok(None);
    }
    let sheet_start = owner
        .sheets
        .get(sheet_index)
        .ok_or_else(|| SourceBackedError::WorksheetNotFound(sheet_index.to_string()))?
        .start;
    let refs = owner
        .workbook_path
        .iter()
        .map(String::as_str)
        .collect::<Vec<_>>();
    let found = owner
        .cfb
        .with_stream_session_at(&refs, sheet_start, |session| {
            session_outcome(query_cell_in_session(
                session,
                owner,
                sheet_index,
                row,
                column,
                execution,
            ))
        })?;
    finish_query(owner, found, execution)
}

fn query_cell_in_session(
    session: &mut SharedOleStreamSession<'_>,
    owner: &Arc<SourceInner>,
    sheet_index: usize,
    row: u32,
    column: u32,
    execution: Option<&ExecutionContext>,
) -> Result<Option<SourceBackedCell>> {
    let sheet = owner
        .sheets
        .get(sheet_index)
        .ok_or_else(|| SourceBackedError::WorksheetNotFound(sheet_index.to_string()))?;
    let mut found = None;
    let mut pending_formula = None;
    let target_row = row as u16;
    let target_column = column as u16;
    let sheet_format = &owner.formatting;
    let mut scan = WorksheetScan::new(sheet.start, sheet.end, owner.limits, execution);
    let mut first = true;
    loop {
        let frame = scan.next_frame(session)?;
        if first {
            first = false;
            if frame.kind != BOF {
                return Err(SourceBackedError::InvalidData(
                    "BoundSheet8 position does not point to a BIFF worksheet BOF".into(),
                ));
            }
            let payload = scan.read_payload(session, &frame)?;
            validate_worksheet_bof(&payload)?;
            continue;
        }
        if frame.kind == EOF {
            if frame.payload_len != 0 {
                return Err(SourceBackedError::InvalidData(
                    "BIFF worksheet EOF has a non-empty payload".into(),
                ));
            }
            if pending_formula.is_some() {
                return Err(SourceBackedError::InvalidData(
                    "string-valued FORMULA lacks STRING result".into(),
                ));
            }
            if let Some(context) = execution {
                context.check().map_err(SourceBackedError::from)?;
            }
            return Ok(found);
        }
        if let Some(mut formula) = pending_formula.take() {
            if frame.kind == STRING {
                let payload = scan.read_payload(session, &frame)?;
                let mut continues = Vec::new();
                let text = loop {
                    match crate::utils::decode_string_record(&payload, &continues)
                        .map_err(SourceBackedError::Parse)?
                    {
                        crate::utils::StringRecordDecode::Complete(text) => break text,
                        crate::utils::StringRecordDecode::NeedContinue => {
                            let next = scan.next_frame(session)?;
                            if next.kind != CONTINUE {
                                return Err(SourceBackedError::InvalidData(
                                    "FORMULA string result continuation is not CONTINUE".into(),
                                ));
                            }
                            continues.push(scan.read_payload(session, &next)?);
                        },
                    }
                };
                if let CellRecord::Formula { value, .. } = &mut formula {
                    *value = FormulaValue::String(text);
                }
                process_cell(
                    &formula,
                    owner,
                    sheet_format,
                    target_row,
                    target_column,
                    &mut found,
                )?;
                continue;
            }
            if !matches!(frame.kind, 0x0221 | 0x0236 | 0x04BC | 0x0091) {
                return Err(SourceBackedError::InvalidData(
                    "string-valued FORMULA is not followed by STRING".into(),
                ));
            }
            pending_formula = Some(formula);
        }

        match frame.kind {
            0x0006 => {
                let payload = scan.read_payload(session, &frame)?;
                let cell = CellRecord::parse(frame.kind, &payload, &owner.encoding)
                    .map_err(SourceBackedError::Parse)?;
                if matches!(
                    cell,
                    CellRecord::Formula {
                        value: FormulaValue::StringPending,
                        ..
                    }
                ) {
                    if cell.row() == target_row && cell.col() == target_column {
                        pending_formula = Some(cell);
                    } else {
                        process_cell(
                            &cell,
                            owner,
                            sheet_format,
                            target_row,
                            target_column,
                            &mut found,
                        )?;
                    }
                } else {
                    process_cell(
                        &cell,
                        owner,
                        sheet_format,
                        target_row,
                        target_column,
                        &mut found,
                    )?;
                }
            },
            0x0201 | 0x0203 | 0x0204 | 0x0205 | 0x027E | 0x00FD => {
                let payload = scan.read_payload(session, &frame)?;
                let cell = CellRecord::parse(frame.kind, &payload, &owner.encoding)
                    .map_err(SourceBackedError::Parse)?;
                process_cell(
                    &cell,
                    owner,
                    sheet_format,
                    target_row,
                    target_column,
                    &mut found,
                )?;
            },
            0x00BD => {
                let payload = scan.read_payload(session, &frame)?;
                for cell in CellRecord::parse_mul_rk(&payload).map_err(SourceBackedError::Parse)? {
                    process_cell(
                        &cell,
                        owner,
                        sheet_format,
                        target_row,
                        target_column,
                        &mut found,
                    )?;
                }
            },
            0x00BE => {
                let payload = scan.read_payload(session, &frame)?;
                for cell in
                    CellRecord::parse_mul_blank(&payload).map_err(SourceBackedError::Parse)?
                {
                    process_cell(
                        &cell,
                        owner,
                        sheet_format,
                        target_row,
                        target_column,
                        &mut found,
                    )?;
                }
            },
            _ => scan.skip_payload(session, &frame)?,
        }
    }
}

fn finish_query(
    owner: &SourceInner,
    found: Option<SourceBackedCell>,
    execution: Option<&ExecutionContext>,
) -> Result<Option<SourceBackedCell>> {
    if let Some(context) = execution {
        context.check().map_err(SourceBackedError::from)?;
    }
    owner.ensure_current()?;
    Ok(found)
}

fn process_cell(
    record: &CellRecord,
    owner: &SourceInner,
    formatting: &Formatting,
    target_row: u16,
    target_column: u16,
    found: &mut Option<SourceBackedCell>,
) -> Result<()> {
    formatting
        .validate_cell_xf(cell_xf_index(record))
        .map_err(SourceBackedError::Parse)?;
    if record.row() != target_row || record.col() != target_column {
        return Ok(());
    }
    let Some(cell) = Cell::from_record_with_formula_context(
        record,
        Some(owner.strings.as_slice()),
        None,
        Some(formatting),
    ) else {
        return Ok(());
    };
    *found = Some(SourceBackedCell {
        row: u32::from(target_row),
        column: u32::from(target_column),
        value: cell.value().clone(),
    });
    Ok(())
}

fn cell_xf_index(record: &CellRecord) -> u16 {
    match record {
        CellRecord::Blank { xf_index, .. }
        | CellRecord::Number { xf_index, .. }
        | CellRecord::Label { xf_index, .. }
        | CellRecord::BoolErr { xf_index, .. }
        | CellRecord::Rk { xf_index, .. }
        | CellRecord::LabelSst { xf_index, .. }
        | CellRecord::Formula { xf_index, .. } => *xf_index,
    }
}

fn map_biff_error(error: litchi_biff::Error) -> SourceBackedError {
    match error {
        litchi_biff::Error::LimitExceeded {
            resource,
            observed,
            maximum,
        } => SourceBackedError::ResourceLimit {
            resource: match resource {
                litchi_biff::Resource::InputBytes => "worksheet scan bytes",
                litchi_biff::Resource::RecordCount => "worksheet scan records",
                _ => "BIFF records",
            },
            observed,
            maximum,
        },
        other => SourceBackedError::Parse(Error::from(other)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn limits_are_finite_by_default() {
        let limits = SourceBackedLimits::default();
        assert!(limits.max_input_bytes > 0);
        assert!(limits.max_global_bytes > 0);
        assert!(limits.max_worksheet_scan_bytes > 0);
    }
}
