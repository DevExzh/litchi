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
use crate::records::{
    BofRecord, BoundSheetRecord, CellRecord, DimensionsRecord, Encoding, FormulaValue,
    SharedStringScanError, SharedStringSstScan, SheetType, decode_shared_string_entry,
    scan_shared_string_records,
};
use crate::{SheetKind, SheetVisibility, Workbook};
use litchi_biff::{Limits as BiffLimits, RecordRef, Records as BiffRecords};
use litchi_cfb::{OleError, SharedOleFile, SharedOleFileLimits, SharedOleStreamCursor};
#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::sheet::Cell as CellTrait;
use litchi_core::{
    ExecutionContext, ExecutionError, ReadAt, SequentialTextWriter, SourceVersion, TextObjectKind,
    TextOutputError, TextOutputOptions, TextOutputReport,
};
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::io::{self, Cursor, Write};
use std::path::Path;
use std::sync::{Arc, Mutex};

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
const DEFAULT_TEXT_CELLS: usize = 1_000_000;
const DEFAULT_TEXT_BYTES: u64 = 128 * 1024 * 1024;
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
    /// Maximum unique cells retained while projecting one worksheet to text.
    pub max_text_cells: usize,
    /// Maximum owned string/error bytes retained while projecting one worksheet.
    pub max_text_bytes: u64,
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
            max_text_cells: DEFAULT_TEXT_CELLS,
            max_text_bytes: DEFAULT_TEXT_BYTES,
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

    /// Sets the maximum unique cells retained for one source-backed text projection.
    #[must_use]
    pub const fn with_max_text_cells(mut self, value: usize) -> Self {
        self.max_text_cells = value;
        self
    }

    /// Sets the maximum owned string/error bytes retained for one source-backed text projection.
    #[must_use]
    pub const fn with_max_text_bytes(mut self, value: u64) -> Self {
        self.max_text_bytes = value;
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
            || self.max_text_cells == 0
            || self.max_text_bytes == 0
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
    Io(io::Error),
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

impl From<io::Error> for SourceBackedError {
    fn from(error: io::Error) -> Self {
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
    sst: Arc<SharedStringSstScan>,
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

struct SourceCheckedTextSink<'owner, 'output, W: ?Sized> {
    output: &'output mut W,
    owner: &'owner SourceInner,
    execution: Option<&'owner ExecutionContext>,
    failure: Arc<Mutex<Option<SourceBackedError>>>,
}

impl<'owner, 'output, W: ?Sized> SourceCheckedTextSink<'owner, 'output, W> {
    fn record_failure(&self, error: SourceBackedError) {
        let mut failure = self
            .failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if failure.is_none() {
            *failure = Some(error);
        }
    }

    fn check(&self) -> io::Result<()> {
        let result = self
            .execution
            .map_or(Ok(()), |context| {
                context.check().map_err(SourceBackedError::from)
            })
            .and_then(|()| self.owner.ensure_current());
        match result {
            Ok(()) => Ok(()),
            Err(error) => {
                let message = error.to_string();
                self.record_failure(error);
                Err(io::Error::other(message))
            },
        }
    }
}

impl<'owner, 'output, W: Write + ?Sized> Write for SourceCheckedTextSink<'owner, 'output, W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.check()?;
        let result = self.output.write(bytes);
        let _ = self.check();
        result
    }

    fn flush(&mut self) -> io::Result<()> {
        self.check()?;
        let result = self.output.flush();
        let _ = self.check();
        result
    }
}

fn take_source_text_failure(
    failure: &Arc<Mutex<Option<SourceBackedError>>>,
) -> Option<SourceBackedError> {
    failure
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
        .take()
}

#[derive(Default)]
struct FallibleTextCollector {
    bytes: Vec<u8>,
    allocation: Option<SourceBackedError>,
}

impl FallibleTextCollector {
    fn push_terminal_newline(&mut self, current_bytes: u64, max_output_bytes: u64) -> Result<()> {
        let required = current_bytes
            .checked_add(1)
            .ok_or(SourceBackedError::ResourceLimit {
                resource: "text output",
                observed: u64::MAX,
                maximum: max_output_bytes,
            })?;
        if required > max_output_bytes {
            return Err(SourceBackedError::ResourceLimit {
                resource: "text output",
                observed: required,
                maximum: max_output_bytes,
            });
        }
        self.bytes
            .try_reserve(1)
            .map_err(|_error| SourceBackedError::Allocation {
                resource: "source-backed text output",
                requested: 1,
            })?;
        self.bytes.push(b'\n');
        Ok(())
    }

    fn take_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

impl Write for FallibleTextCollector {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if let Err(_error) = self.bytes.try_reserve(bytes.len()) {
            let typed = SourceBackedError::Allocation {
                resource: "source-backed text output",
                requested: bytes.len() as u64,
            };
            let message = typed.to_string();
            self.allocation = Some(typed);
            return Err(io::Error::other(message));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct SourceTextSheet {
    cells: HashMap<(u16, u16), litchi_core::sheet::CellValue>,
    max_row: u16,
    max_col: u16,
    retained_text_bytes: u64,
}

impl SourceTextSheet {
    fn new() -> Self {
        Self {
            cells: HashMap::new(),
            max_row: 0,
            max_col: 0,
            retained_text_bytes: 0,
        }
    }

    fn insert(
        &mut self,
        row: u16,
        column: u16,
        value: litchi_core::sheet::CellValue,
        limits: SourceBackedLimits,
    ) -> Result<()> {
        let old_bytes = self
            .cells
            .get(&(row, column))
            .map(retained_text_bytes)
            .unwrap_or(0);
        let new_bytes = retained_text_bytes(&value);
        let retained = self
            .retained_text_bytes
            .checked_sub(old_bytes)
            .and_then(|bytes| bytes.checked_add(new_bytes))
            .ok_or(SourceBackedError::ResourceLimit {
                resource: "text bytes",
                observed: u64::MAX,
                maximum: limits.max_text_bytes,
            })?;
        if retained > limits.max_text_bytes {
            return Err(SourceBackedError::ResourceLimit {
                resource: "text bytes",
                observed: retained,
                maximum: limits.max_text_bytes,
            });
        }

        let is_new = !self.cells.contains_key(&(row, column));
        if is_new {
            let observed =
                self.cells
                    .len()
                    .checked_add(1)
                    .ok_or(SourceBackedError::ResourceLimit {
                        resource: "text cells",
                        observed: u64::MAX,
                        maximum: limits.max_text_cells as u64,
                    })?;
            if observed > limits.max_text_cells {
                return Err(SourceBackedError::ResourceLimit {
                    resource: "text cells",
                    observed: observed as u64,
                    maximum: limits.max_text_cells as u64,
                });
            }
            self.cells
                .try_reserve(1)
                .map_err(|_error| SourceBackedError::Allocation {
                    resource: "source-backed text cells",
                    requested: 1,
                })?;
        }

        let _ = self.cells.insert((row, column), value);
        self.retained_text_bytes = retained;
        self.max_row = self.max_row.max(row);
        self.max_col = self.max_col.max(column);
        Ok(())
    }
}

fn retained_text_bytes(value: &litchi_core::sheet::CellValue) -> u64 {
    match value {
        litchi_core::sheet::CellValue::String(value)
        | litchi_core::sheet::CellValue::Error(value) => {
            u64::try_from(value.len()).unwrap_or(u64::MAX)
        },
        litchi_core::sheet::CellValue::Formula {
            formula,
            cached_value,
            ..
        } => u64::try_from(formula.len())
            .unwrap_or(u64::MAX)
            .saturating_add(
                cached_value
                    .as_deref()
                    .map(retained_text_bytes)
                    .unwrap_or(0),
            ),
        litchi_core::sheet::CellValue::Empty
        | litchi_core::sheet::CellValue::Bool(_)
        | litchi_core::sheet::CellValue::Int(_)
        | litchi_core::sheet::CellValue::Float(_)
        | litchi_core::sheet::CellValue::DateTime(_) => 0,
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

    /// Extracts all ordinary worksheet text through the bounded source-backed
    /// projection. The returned string retains the legacy trailing newline.
    pub fn text(&self) -> Result<String> {
        self.text_impl(None)
    }

    /// Extracts all ordinary worksheet text with cooperative cancellation.
    pub fn text_with_execution(&self, execution: &ExecutionContext) -> Result<String> {
        self.text_impl(Some(execution))
    }

    /// Streams ordinary worksheet rows to a caller-owned sink.
    ///
    /// Each logical row is one paragraph-like object. Cells are emitted in a
    /// dense rectangular range with tab separators. The standard text-output
    /// policy controls object separators, empty rows, and output limits; this
    /// method does not append a terminal separator.
    pub fn write_text_to<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: TextOutputOptions<'_>,
    ) -> std::result::Result<TextOutputReport, TextOutputError<SourceBackedError>> {
        self.write_text_to_impl(output, options, None)
    }

    /// Streams ordinary worksheet rows with cooperative cancellation.
    pub fn write_text_to_with_execution<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: TextOutputOptions<'_>,
        execution: &ExecutionContext,
    ) -> std::result::Result<TextOutputReport, TextOutputError<SourceBackedError>> {
        self.write_text_to_impl(output, options, Some(execution))
    }

    fn text_impl(&self, execution: Option<&ExecutionContext>) -> Result<String> {
        let mut collector = FallibleTextCollector::default();
        let report = match self.write_text_to_impl(
            &mut collector,
            TextOutputOptions::default(),
            execution,
        ) {
            Ok(report) => report,
            Err(error) => {
                let allocation = collector.allocation.take();
                return Err(map_text_output_error(error, allocation));
            },
        };
        if report.objects_written() != 0 {
            check_text_state(&self.inner, execution)?;
            collector.push_terminal_newline(
                report.bytes_written(),
                TextOutputOptions::default().max_output_bytes(),
            )?;
            check_text_state(&self.inner, execution)?;
        }
        self.inner.ensure_current()?;
        String::from_utf8(collector.take_bytes())
            .map_err(|_error| SourceBackedError::InvalidData("text output was not UTF-8".into()))
    }

    fn write_text_to_impl<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: TextOutputOptions<'_>,
        execution: Option<&ExecutionContext>,
    ) -> std::result::Result<TextOutputReport, TextOutputError<SourceBackedError>> {
        let failure = Arc::new(Mutex::new(None));
        let mut checked_output = SourceCheckedTextSink {
            output,
            owner: &self.inner,
            execution,
            failure: Arc::clone(&failure),
        };
        let mut writer = SequentialTextWriter::new(&mut checked_output, options);
        let conversion = (|| {
            check_text_state(&self.inner, execution)
                .map_err(|source| writer.document_error(source))?;
            let refs = self
                .inner
                .workbook_path
                .iter()
                .map(String::as_str)
                .collect::<Vec<_>>();
            for sheet in self
                .inner
                .sheets
                .iter()
                .filter(|sheet| sheet.worksheet_index.is_some())
            {
                check_text_state(&self.inner, execution)
                    .map_err(|source| writer.document_error(source))?;
                let collected = scan_text_sheet(&self.inner, sheet, &refs, execution)
                    .map_err(|source| writer.document_error(source))?;
                check_text_state(&self.inner, execution)
                    .map_err(|source| writer.document_error(source))?;
                write_text_sheet(&self.inner, &collected, &mut writer, execution)?;
            }
            Ok::<(), TextOutputError<SourceBackedError>>(())
        })();

        let progress = writer.progress();
        let source = take_source_text_failure(&failure)
            .or_else(|| execution.and_then(|context| context.check().err().map(Into::into)))
            .or_else(|| self.inner.ensure_current().err());
        if let Some(source) = source {
            return Err(TextOutputError::Document { source, progress });
        }
        conversion.map(|()| writer.finish())
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
            sst: Arc::new(parsed.sst),
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
    sst: SharedStringSstScan,
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
    let mut offset = 0_u64;
    let mut record_count = 0_usize;
    let mut cursor = cfb
        .stream_cursor_at(&refs, 0)
        .map_err(SourceBackedError::from)?;

    // The first pass reads only BIFF headers.  Besides avoiding payload reads
    // for encrypted workbooks, this establishes the exact global boundary so
    // the semantic pass below can issue one bounded logical range read.
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
        cursor
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
        cursor
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
                sst_refs
                    .try_reserve(1)
                    .map_err(|_error| SourceBackedError::Allocation {
                        resource: "SST record references",
                        requested: 1,
                    })?;
                sst_refs.push(record);
                while records
                    .get(i + 1)
                    .is_some_and(|next| next.kind().get() == CONTINUE)
                {
                    i += 1;
                    sst_refs
                        .try_reserve(1)
                        .map_err(|_error| SourceBackedError::Allocation {
                            resource: "SST record references",
                            requested: 1,
                        })?;
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
    let sst = scan_shared_string_records(&sst_refs).map_err(map_shared_string_error)?;

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
        sst,
        formatting,
        encoding,
    })
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
    cursor: SharedOleStreamCursor<'a>,
    upper_bound: u64,
    scanned_bytes: u64,
    scanned_records: usize,
    limits: SourceBackedLimits,
    execution: Option<&'a ExecutionContext>,
    scratch: Vec<u8>,
}

impl<'a> WorksheetScan<'a> {
    fn new(
        cfb: &'a SharedOleFile,
        path: &'a [&'a str],
        start: u64,
        upper_bound: u64,
        limits: SourceBackedLimits,
        execution: Option<&'a ExecutionContext>,
    ) -> Result<Self> {
        let cursor = cfb
            .stream_cursor_at(path, start)
            .map_err(SourceBackedError::from)?;
        Ok(Self {
            cursor,
            upper_bound,
            scanned_bytes: 0,
            scanned_records: 0,
            limits,
            execution,
            scratch: Vec::new(),
        })
    }

    fn next_frame(&mut self) -> Result<WorksheetFrame> {
        self.check_execution()?;
        let cursor_position = self.cursor.position();
        if cursor_position
            .checked_add(4)
            .is_none_or(|value| value > self.upper_bound)
        {
            return Err(SourceBackedError::InvalidData(
                "BIFF worksheet has no complete record header before its boundary".into(),
            ));
        }
        let mut header = [0_u8; 4];
        self.cursor
            .read_exact(&mut header)
            .map_err(SourceBackedError::from)?;
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

    fn read_payload(&mut self, frame: &WorksheetFrame) -> Result<&[u8]> {
        self.check_execution()?;
        self.scratch.clear();
        if self.scratch.capacity() < frame.payload_len {
            self.scratch
                .try_reserve_exact(frame.payload_len)
                .map_err(|_error| SourceBackedError::Allocation {
                    resource: "source-backed worksheet payload",
                    requested: frame.payload_len as u64,
                })?;
        }
        self.scratch.resize(frame.payload_len, 0);
        self.cursor
            .read_exact(&mut self.scratch)
            .map_err(SourceBackedError::from)?;
        Ok(&self.scratch)
    }

    fn take_payload(&mut self, frame: &WorksheetFrame) -> Result<Vec<u8>> {
        self.read_payload(frame)?;
        Ok(std::mem::take(&mut self.scratch))
    }

    fn recycle_payload(&mut self, payload: Vec<u8>) {
        self.scratch = payload;
    }

    fn skip_payload(&mut self, frame: &WorksheetFrame) -> Result<()> {
        self.check_execution()?;
        self.cursor
            .skip_forward(frame.payload_len as u64)
            .map_err(SourceBackedError::from)
    }

    fn check_execution(&self) -> Result<()> {
        self.execution.map_or(Ok(()), |context| {
            context.check().map_err(SourceBackedError::from)
        })
    }
}

fn check_text_state(owner: &SourceInner, execution: Option<&ExecutionContext>) -> Result<()> {
    if let Some(context) = execution {
        context.check().map_err(SourceBackedError::from)?;
    }
    owner.ensure_current()
}

fn map_text_output_error(
    error: TextOutputError<SourceBackedError>,
    collector_failure: Option<SourceBackedError>,
) -> SourceBackedError {
    match error {
        TextOutputError::Document { source, .. } => source,
        TextOutputError::Limit { limit, .. } => SourceBackedError::ResourceLimit {
            resource: "text output",
            observed: limit.observed(),
            maximum: limit.limit(),
        },
        TextOutputError::Sink { source, .. } => {
            collector_failure.unwrap_or(SourceBackedError::Io(source))
        },
        TextOutputError::NonDeterministicFragments { .. } => SourceBackedError::InvalidData(
            "source-backed XLS text fragments were not deterministic".into(),
        ),
        _ => SourceBackedError::InvalidData("unsupported text output error".into()),
    }
}

fn scan_text_sheet(
    owner: &SourceInner,
    sheet: &SheetEntry,
    refs: &[&str],
    execution: Option<&ExecutionContext>,
) -> Result<SourceTextSheet> {
    let mut collected = SourceTextSheet::new();
    let mut scan = WorksheetScan::new(
        &owner.cfb,
        refs,
        sheet.start,
        sheet.end,
        owner.limits,
        execution,
    )?;
    let mut pending_formula = None;
    let mut first = true;

    loop {
        let frame = scan.next_frame()?;
        if first {
            first = false;
            if frame.kind != BOF {
                return Err(SourceBackedError::InvalidData(
                    "BoundSheet8 position does not point to a BIFF worksheet BOF".into(),
                ));
            }
            let payload = scan.read_payload(&frame)?;
            validate_worksheet_bof(payload)?;
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
            return Ok(collected);
        }

        if let Some(mut formula) = pending_formula.take() {
            if frame.kind == STRING {
                let payload = scan.take_payload(&frame)?;
                let mut continues = Vec::new();
                let text = loop {
                    match crate::utils::decode_string_record(&payload, &continues)
                        .map_err(SourceBackedError::Parse)?
                    {
                        crate::utils::StringRecordDecode::Complete(text) => break text,
                        crate::utils::StringRecordDecode::NeedContinue => {
                            let next = scan.next_frame()?;
                            if next.kind != CONTINUE {
                                return Err(SourceBackedError::InvalidData(
                                    "FORMULA string result continuation is not CONTINUE".into(),
                                ));
                            }
                            continues.try_reserve(1).map_err(|_error| {
                                SourceBackedError::Allocation {
                                    resource: "formula STRING continuations",
                                    requested: 1,
                                }
                            })?;
                            continues.push(scan.take_payload(&next)?);
                        },
                    }
                };
                scan.recycle_payload(payload);
                if let CellRecord::Formula { value, .. } = &mut formula {
                    *value = FormulaValue::String(text);
                }
                collect_source_cell(&formula, owner, &mut collected, execution)?;
                continue;
            }
            if !matches!(frame.kind, 0x0221 | 0x0236 | 0x04BC | 0x0091) {
                return Err(SourceBackedError::InvalidData(
                    "string-valued FORMULA is not followed by STRING".into(),
                ));
            }
            pending_formula = Some(formula);
        }

        if frame.kind == STRING {
            return Err(SourceBackedError::InvalidData(
                "STRING record has no pending FORMULA result".into(),
            ));
        }

        match frame.kind {
            0x0200 => {
                let payload = scan.read_payload(&frame)?;
                if let Ok(dimensions) = DimensionsRecord::parse(payload) {
                    let max_row = dimensions.last_row.saturating_sub(1);
                    let max_column = dimensions.last_col.saturating_sub(1);
                    collected.max_row = collected
                        .max_row
                        .max(u16::try_from(max_row).unwrap_or(u16::MAX));
                    collected.max_col = collected
                        .max_col
                        .max(u16::try_from(max_column).unwrap_or(u16::MAX));
                }
            },
            0x0006 => {
                let payload = scan.read_payload(&frame)?;
                let cell = CellRecord::parse(frame.kind, payload, &owner.encoding)
                    .map_err(SourceBackedError::Parse)?;
                if matches!(
                    cell,
                    CellRecord::Formula {
                        value: FormulaValue::StringPending,
                        ..
                    }
                ) {
                    pending_formula = Some(cell);
                } else {
                    collect_source_cell(&cell, owner, &mut collected, execution)?;
                }
            },
            0x0201 | 0x0203 | 0x0204 | 0x0205 | 0x027E | 0x00FD => {
                let payload = scan.read_payload(&frame)?;
                let cell = CellRecord::parse(frame.kind, payload, &owner.encoding)
                    .map_err(SourceBackedError::Parse)?;
                collect_source_cell(&cell, owner, &mut collected, execution)?;
            },
            0x00BD => {
                let payload = scan.read_payload(&frame)?;
                let mut processing = Ok(());
                CellRecord::visit_mul_rk(payload, |cell| {
                    if processing.is_ok() {
                        processing = collect_source_cell(&cell, owner, &mut collected, execution);
                    }
                })
                .map_err(SourceBackedError::Parse)?;
                processing?;
            },
            0x00BE => {
                let payload = scan.read_payload(&frame)?;
                let mut processing = Ok(());
                CellRecord::visit_mul_blank(payload, |cell| {
                    if processing.is_ok() {
                        processing = collect_source_cell(&cell, owner, &mut collected, execution);
                    }
                })
                .map_err(SourceBackedError::Parse)?;
                processing?;
            },
            _ => scan.skip_payload(&frame)?,
        }
    }
}

fn write_text_sheet<'options, 'output, W: Write + ?Sized>(
    owner: &SourceInner,
    sheet: &SourceTextSheet,
    writer: &mut SequentialTextWriter<'options, 'output, W>,
    execution: Option<&ExecutionContext>,
) -> std::result::Result<(), TextOutputError<SourceBackedError>> {
    let mut row = 0_u16;
    loop {
        check_text_state(owner, execution).map_err(|source| writer.document_error(source))?;
        let mut value = String::new();
        let mut column = 0_u16;
        loop {
            if column != 0 {
                value.try_reserve(1).map_err(|_error| {
                    writer.document_error(SourceBackedError::Allocation {
                        resource: "source-backed text row",
                        requested: 1,
                    })
                })?;
                value.push('\t');
            }
            if let Some(cell) = sheet.cells.get(&(row, column)) {
                append_source_cell_text(&mut value, cell)
                    .map_err(|source| writer.document_error(source))?;
            }
            if column == sheet.max_col {
                break;
            }
            column = column.checked_add(1).ok_or_else(|| {
                writer.document_error(SourceBackedError::InvalidData(
                    "source-backed XLS text column overflow".into(),
                ))
            })?;
        }
        writer.write_object(TextObjectKind::Paragraph, &value)?;
        if row == sheet.max_row {
            break;
        }
        row = row.checked_add(1).ok_or_else(|| {
            writer.document_error(SourceBackedError::InvalidData(
                "source-backed XLS text row overflow".into(),
            ))
        })?;
    }
    Ok(())
}

struct TextByteCounter {
    bytes: usize,
}

impl fmt::Write for TextByteCounter {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.bytes = self.bytes.checked_add(value.len()).ok_or(fmt::Error)?;
        Ok(())
    }
}

fn append_counted_text<F>(output: &mut String, render: F) -> Result<()>
where
    F: Fn(&mut dyn fmt::Write) -> fmt::Result,
{
    let mut counter = TextByteCounter { bytes: 0 };
    render(&mut counter).map_err(|_error| {
        SourceBackedError::InvalidData("text formatting length overflow".into())
    })?;
    let additional = u64::try_from(counter.bytes).unwrap_or(u64::MAX);
    let current = u64::try_from(output.len()).unwrap_or(u64::MAX);
    current
        .checked_add(additional)
        .ok_or(SourceBackedError::Allocation {
            resource: "source-backed text row",
            requested: u64::MAX,
        })?;
    output
        .try_reserve(counter.bytes)
        .map_err(|_error| SourceBackedError::Allocation {
            resource: "source-backed text row",
            requested: counter.bytes as u64,
        })?;
    render(output).map_err(|_error| SourceBackedError::InvalidData("text formatting failed".into()))
}

fn append_source_cell_text(
    output: &mut String,
    value: &litchi_core::sheet::CellValue,
) -> Result<()> {
    match value {
        litchi_core::sheet::CellValue::Empty => Ok(()),
        litchi_core::sheet::CellValue::Bool(value) => append_counted_text(output, |writer| {
            fmt::Write::write_str(writer, if *value { "TRUE" } else { "FALSE" })
        }),
        litchi_core::sheet::CellValue::Int(value) => {
            append_counted_text(output, |writer| fmt::write(writer, format_args!("{value}")))
        },
        litchi_core::sheet::CellValue::Float(value)
        | litchi_core::sheet::CellValue::DateTime(value) => {
            append_counted_text(output, |writer| fmt::write(writer, format_args!("{value}")))
        },
        litchi_core::sheet::CellValue::String(value)
        | litchi_core::sheet::CellValue::Error(value) => {
            append_counted_text(output, |writer| fmt::Write::write_str(writer, value))
        },
        litchi_core::sheet::CellValue::Formula {
            formula,
            cached_value,
            ..
        } => match cached_value.as_deref() {
            Some(litchi_core::sheet::CellValue::Empty) | None => {
                append_counted_text(output, |writer| {
                    fmt::Write::write_char(writer, '=')?;
                    fmt::Write::write_str(writer, formula)
                })
            },
            Some(value) => append_source_cell_text(output, value),
        },
    }
}

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
    let sheet = owner
        .sheets
        .get(sheet_index)
        .ok_or_else(|| SourceBackedError::WorksheetNotFound(sheet_index.to_string()))?;
    let refs = owner
        .workbook_path
        .iter()
        .map(String::as_str)
        .collect::<Vec<_>>();
    let mut found = None;
    let mut pending_formula = None;
    let target_row = row as u16;
    let target_column = column as u16;
    let sheet_format = &owner.formatting;
    let mut scan = WorksheetScan::new(
        &owner.cfb,
        &refs,
        sheet.start,
        sheet.end,
        owner.limits,
        execution,
    )?;
    let mut first = true;
    loop {
        let frame = scan.next_frame()?;
        if first {
            first = false;
            if frame.kind != BOF {
                return Err(SourceBackedError::InvalidData(
                    "BoundSheet8 position does not point to a BIFF worksheet BOF".into(),
                ));
            }
            let payload = scan.read_payload(&frame)?;
            validate_worksheet_bof(payload)?;
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
            return finish_query(owner, found, execution);
        }
        if let Some(mut formula) = pending_formula.take() {
            if frame.kind == STRING {
                let payload = scan.take_payload(&frame)?;
                let mut continues = Vec::new();
                let text = loop {
                    match crate::utils::decode_string_record(&payload, &continues)
                        .map_err(SourceBackedError::Parse)?
                    {
                        crate::utils::StringRecordDecode::Complete(text) => break text,
                        crate::utils::StringRecordDecode::NeedContinue => {
                            let next = scan.next_frame()?;
                            if next.kind != CONTINUE {
                                return Err(SourceBackedError::InvalidData(
                                    "FORMULA string result continuation is not CONTINUE".into(),
                                ));
                            }
                            continues.push(scan.take_payload(&next)?);
                        },
                    }
                };
                scan.recycle_payload(payload);
                if let CellRecord::Formula { value, .. } = &mut formula {
                    *value = FormulaValue::String(text);
                }
                process_cell(
                    &formula,
                    owner,
                    sheet_format,
                    target_row,
                    target_column,
                    execution,
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
                let payload = scan.read_payload(&frame)?;
                let cell = CellRecord::parse(frame.kind, payload, &owner.encoding)
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
                            execution,
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
                        execution,
                        &mut found,
                    )?;
                }
            },
            0x0201 | 0x0203 | 0x0204 | 0x0205 | 0x027E | 0x00FD => {
                let payload = scan.read_payload(&frame)?;
                let cell = CellRecord::parse(frame.kind, payload, &owner.encoding)
                    .map_err(SourceBackedError::Parse)?;
                process_cell(
                    &cell,
                    owner,
                    sheet_format,
                    target_row,
                    target_column,
                    execution,
                    &mut found,
                )?;
            },
            0x00BD => {
                let payload = scan.read_payload(&frame)?;
                let mut processing = Ok(());
                CellRecord::visit_mul_rk(payload, |cell| {
                    if processing.is_ok() {
                        processing = process_cell(
                            &cell,
                            owner,
                            sheet_format,
                            target_row,
                            target_column,
                            execution,
                            &mut found,
                        );
                    }
                })
                .map_err(SourceBackedError::Parse)?;
                processing?;
            },
            0x00BE => {
                let payload = scan.read_payload(&frame)?;
                let mut processing = Ok(());
                CellRecord::visit_mul_blank(payload, |cell| {
                    if processing.is_ok() {
                        processing = process_cell(
                            &cell,
                            owner,
                            sheet_format,
                            target_row,
                            target_column,
                            execution,
                            &mut found,
                        );
                    }
                })
                .map_err(SourceBackedError::Parse)?;
                processing?;
            },
            _ => scan.skip_payload(&frame)?,
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

fn resolve_shared_string(
    owner: &SourceInner,
    string_index: u32,
    execution: Option<&ExecutionContext>,
) -> Result<litchi_core::sheet::CellValue> {
    if owner.sst.segments.is_empty() {
        return Ok(litchi_core::sheet::CellValue::Error(
            "SST not available".to_owned(),
        ));
    }
    let Some(index) = usize::try_from(string_index).ok() else {
        return Ok(litchi_core::sheet::CellValue::Error(format!(
            "Invalid SST index: {string_index} (max: {})",
            owner.sst.entries.len()
        )));
    };
    let Some(location) = owner.sst.entries.get(index).copied() else {
        return Ok(litchi_core::sheet::CellValue::Error(format!(
            "Invalid SST index: {string_index} (max: {})",
            owner.sst.entries.len()
        )));
    };
    if location.start >= location.end {
        return Err(SourceBackedError::InvalidData(
            "SST entry locator has an empty span".to_owned(),
        ));
    }

    if let Some(context) = execution {
        context.check().map_err(SourceBackedError::from)?;
    }
    owner.ensure_current()?;

    let mut first_segment = None;
    for (segment_index, segment) in owner.sst.segments.iter().enumerate() {
        let segment_end = segment
            .logical_offset
            .checked_add(segment.len)
            .ok_or_else(|| SourceBackedError::InvalidData("SST segment span overflow".into()))?;
        if location.start >= segment.logical_offset && location.start < segment_end {
            first_segment = Some(segment_index);
            break;
        }
    }
    let Some(first_segment) = first_segment else {
        return Err(SourceBackedError::InvalidData(
            "SST entry locator is outside its segments".to_owned(),
        ));
    };

    let first = &owner.sst.segments[first_segment];
    let first_offset = first
        .source_offset
        .checked_add((location.start - first.logical_offset) as u64)
        .ok_or_else(|| SourceBackedError::InvalidData("SST source offset overflow".into()))?;
    let refs = owner
        .workbook_path
        .iter()
        .map(String::as_str)
        .collect::<Vec<_>>();
    let mut cursor = owner
        .cfb
        .stream_cursor_at(&refs, first_offset)
        .map_err(SourceBackedError::from)?;
    let mut chunks = Vec::<Vec<u8>>::new();
    chunks
        .try_reserve_exact(owner.sst.segments.len().saturating_sub(first_segment))
        .map_err(|_| SourceBackedError::Allocation {
            resource: "selected SST chunks",
            requested: owner.sst.segments.len().saturating_sub(first_segment) as u64,
        })?;

    for segment in owner.sst.segments.iter().skip(first_segment) {
        let segment_end = segment
            .logical_offset
            .checked_add(segment.len)
            .ok_or_else(|| SourceBackedError::InvalidData("SST segment span overflow".into()))?;
        let start = location.start.max(segment.logical_offset);
        let end = location.end.min(segment_end);
        if start >= end {
            if segment.logical_offset >= location.end {
                break;
            }
            continue;
        }
        if let Some(context) = execution {
            context.check().map_err(SourceBackedError::from)?;
        }
        owner.ensure_current()?;
        let source_offset = segment
            .source_offset
            .checked_add((start - segment.logical_offset) as u64)
            .ok_or_else(|| SourceBackedError::InvalidData("SST source offset overflow".into()))?;
        cursor
            .skip_to(source_offset)
            .map_err(SourceBackedError::from)?;
        let length = end - start;
        let mut chunk = Vec::new();
        chunk
            .try_reserve_exact(length)
            .map_err(|_| SourceBackedError::Allocation {
                resource: "selected SST entry",
                requested: length as u64,
            })?;
        chunk.resize(length, 0);
        cursor
            .read_exact(&mut chunk)
            .map_err(SourceBackedError::from)?;
        chunks.push(chunk);
    }

    let mut slices = Vec::new();
    slices
        .try_reserve_exact(chunks.len())
        .map_err(|_| SourceBackedError::Allocation {
            resource: "selected SST parser segments",
            requested: chunks.len() as u64,
        })?;
    for chunk in &chunks {
        slices.push(chunk.as_slice());
    }
    let decoded = decode_shared_string_entry(&slices);
    match decoded {
        Ok(value) => {
            owner.ensure_current()?;
            Ok(litchi_core::sheet::CellValue::String(value))
        },
        Err(error) => {
            owner.ensure_current()?;
            Err(map_shared_string_error(error))
        },
    }
}

fn decode_source_cell(
    record: &CellRecord,
    owner: &SourceInner,
    formatting: &Formatting,
    execution: Option<&ExecutionContext>,
) -> Result<Option<(u16, u16, litchi_core::sheet::CellValue)>> {
    formatting
        .validate_cell_xf(cell_xf_index(record))
        .map_err(SourceBackedError::Parse)?;
    let Some(cell) = Cell::from_record_with_formula_context(record, None, None, Some(formatting))
    else {
        return Ok(None);
    };
    let value = if let Some(string_index) = cell.shared_string_index() {
        if owner.sst.segments.is_empty() {
            litchi_core::sheet::CellValue::Error(format!(
                "Invalid SST index: {string_index} (max: 0)"
            ))
        } else {
            resolve_shared_string(owner, string_index, execution)?
        }
    } else {
        cell.value().clone()
    };
    Ok(Some((record.row(), record.col(), value)))
}

fn collect_source_cell(
    record: &CellRecord,
    owner: &SourceInner,
    collected: &mut SourceTextSheet,
    execution: Option<&ExecutionContext>,
) -> Result<()> {
    if let Some((row, column, value)) =
        decode_source_cell(record, owner, &owner.formatting, execution)?
    {
        collected.insert(row, column, value, owner.limits)?;
    }
    Ok(())
}

fn process_cell(
    record: &CellRecord,
    owner: &SourceInner,
    formatting: &Formatting,
    target_row: u16,
    target_column: u16,
    execution: Option<&ExecutionContext>,
    found: &mut Option<SourceBackedCell>,
) -> Result<()> {
    formatting
        .validate_cell_xf(cell_xf_index(record))
        .map_err(SourceBackedError::Parse)?;
    if record.row() != target_row || record.col() != target_column {
        return Ok(());
    }
    let Some(cell) = Cell::from_record_with_formula_context(record, None, None, Some(formatting))
    else {
        return Ok(());
    };
    let value = if let Some(string_index) = cell.shared_string_index() {
        resolve_shared_string(owner, string_index, execution)?
    } else {
        cell.value().clone()
    };
    *found = Some(SourceBackedCell {
        row: u32::from(target_row),
        column: u32::from(target_column),
        value,
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

fn map_shared_string_error(error: SharedStringScanError) -> SourceBackedError {
    match error {
        SharedStringScanError::Biff(error) => SourceBackedError::Parse(error),
        SharedStringScanError::Invalid(message) => {
            SourceBackedError::Parse(Error::InvalidData(message))
        },
        SharedStringScanError::Allocation {
            resource,
            requested,
        } => SourceBackedError::Allocation {
            resource,
            requested: requested as u64,
        },
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
