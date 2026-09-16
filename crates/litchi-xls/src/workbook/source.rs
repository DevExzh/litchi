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
    MeasuredCell, SharedStringScanError, SharedStringSstScan, SheetType,
    decode_shared_string_entry, scan_shared_string_records,
};
use crate::{SheetKind, SheetVisibility, Workbook};
use litchi_biff::{Limits as BiffLimits, RecordRef, Records as BiffRecords};
use litchi_cfb::{
    OleError, SharedOleFile, SharedOleFileLimits, SharedOleStreamCursor, StreamChainHint,
};
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
/// Workbook globals records whose bytes are fetched exactly rather than in a
/// window: one read for the header, one for the payload and the next record's
/// header.
///
/// [MS-XLS] 2.1.7.20.1 gives the globals substream as
/// `GLOBALS = BOF [WriteProtect] [FilePass] [Template] ...`, so a `FilePass`
/// record is one of the first three. Four records cover that prefix, and the
/// last one's fetch also buffers the fifth record's header.
const GLOBALS_EXACT_PROLOGUE_RECORDS: usize = 4;
/// First window fill after the exact prologue; one CFB sector.
const GLOBALS_FIRST_WINDOW_BYTES: u64 = 512;
/// Upper bound the window fill size doubles to.
const GLOBALS_MAX_WINDOW_BYTES: u64 = 64 * 1024;
/// First window fill of a worksheet scan; one CFB sector.
const WORKSHEET_FIRST_WINDOW_BYTES: u64 = 512;
/// Upper bound the worksheet window fill size doubles to.
const WORKSHEET_MAX_WINDOW_BYTES: u64 = 64 * 1024;
/// Mean framed bytes per record at or below which a worksheet scan fills in
/// windows; above it each fill covers only the bytes the current record needs.
const WORKSHEET_DENSE_FRAME_BYTES: u64 = 1024;

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
        // `let ... else` rather than `ok_or`: the limit value is built only on
        // the path that returns it. `ok_or` builds it, and drops it again, on
        // every cell of every text extraction, and clippy's
        // `unnecessary_lazy_evaluations` rejects the `ok_or_else` spelling.
        let Some(retained) = self
            .retained_text_bytes
            .checked_sub(old_bytes)
            .and_then(|bytes| bytes.checked_add(new_bytes))
        else {
            return Err(SourceBackedError::ResourceLimit {
                resource: "text bytes",
                observed: u64::MAX,
                maximum: limits.max_text_bytes,
            });
        };
        if retained > limits.max_text_bytes {
            return Err(SourceBackedError::ResourceLimit {
                resource: "text bytes",
                observed: retained,
                maximum: limits.max_text_bytes,
            });
        }

        let is_new = !self.cells.contains_key(&(row, column));
        if is_new {
            let Some(observed) = self.cells.len().checked_add(1) else {
                return Err(SourceBackedError::ResourceLimit {
                    resource: "text cells",
                    observed: u64::MAX,
                    maximum: limits.max_text_cells as u64,
                });
            };
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
            // One shared-string resolver for the whole document. Its chain
            // position carries across sheets because every sheet resolves out
            // of the same string table; a position that does not apply to a
            // resolve is discarded by the reader, which then walks cold.
            let mut strings = SharedStringResolver::new(&self.inner, &refs);
            // One worksheet-region chain position for the whole document,
            // disjoint from the resolver's. `BoundSheet8` positions ascend, so
            // each sheet's cursor resumes from the previous sheet's start
            // instead of re-walking the chain from the stream's first sector;
            // a position that does not apply is discarded by the reader, which
            // then walks cold exactly as before.
            let mut sheet_chain = self.inner.cfb.chain_hint();
            for sheet in self
                .inner
                .sheets
                .iter()
                .filter(|sheet| sheet.worksheet_index.is_some())
            {
                check_text_state(&self.inner, execution)
                    .map_err(|source| writer.document_error(source))?;
                let collected = scan_text_sheet(
                    &self.inner,
                    sheet,
                    &refs,
                    execution,
                    &mut strings,
                    &mut sheet_chain,
                )
                .map_err(|source| writer.document_error(source))?;
                check_text_state(&self.inner, execution)
                    .map_err(|source| writer.document_error(source))?;
                write_text_sheet(&collected, &mut writer, execution)?;
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
        // `SharedOleFile::source_version` already observes the source this view
        // was opened over and refuses when it has moved, and `source` is that
        // same object (`source_arc`). An `ensure_current_parts` call here would
        // observe it a second time with nothing in between -- no source byte is
        // consumed, and the captured-version equality it also checks is this
        // very value -- so it would re-prove what this line proved. Change 0560
        // collapsed the same pair inside `ensure_current_parts` itself.
        let expected_version = cfb.source_version().map_err(SourceBackedError::from)?;
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

    /// Reads retained in-memory workbook metadata behind one trailing fence.
    ///
    /// `operation` reads only values this snapshot already owns; it consumes no
    /// source bytes. One observation after it therefore proves everything two
    /// observations around it proved, and it is the trailing one that bounds
    /// what the caller receives.
    fn metadata<T>(&self, operation: impl FnOnce(&SourceInner) -> T) -> Result<T> {
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

/// Fences the retained source once for both the workbook snapshot and the CFB
/// view it was opened over.
///
/// `SourceInner` takes its source from `SharedOleFile::source_arc` and its
/// expected version from the same view, so the two expectations are the same
/// value observed against the same object. One observation therefore discharges
/// both, and `SharedOleFile::source_version` would only repeat it: nothing
/// between the observations reads a source byte. The debug assertion records
/// the identity this relies on.
fn ensure_current_parts(
    source: &Arc<dyn ReadAt>,
    cfb: &SharedOleFile,
    expected_version: SourceVersion,
) -> Result<()> {
    debug_assert_eq!(
        cfb.captured_source_version(),
        expected_version,
        "the workbook snapshot and its CFB view must share one captured source version"
    );
    let observed = source.version().map_err(SourceBackedError::Io)?;
    if observed != expected_version || cfb.captured_source_version() != expected_version {
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

    /// Reports every stored cell of this worksheet, in stream order, from
    /// **one** validated scan.
    ///
    /// [`cell`](Self::cell) scans the whole worksheet substream to its EOF for
    /// every lookup, so reading a worksheet cell by cell costs one full
    /// validated scan per cell. This walks the same records once, under the
    /// same limits, taking the same checks in the same order, and hands each
    /// stored cell to `visitor`.
    ///
    /// The scan always runs to the worksheet's EOF: the visitor cannot stop it
    /// early, so a malformed record beyond the last cell is refused here
    /// exactly as [`cell`](Self::cell) refuses it. A visitor that returns an
    /// error does end the scan, and that error is returned unchanged.
    ///
    /// A cell is *stored* when the worksheet holds a record for it; blank and
    /// unformatted-blank records are reported, positions with no record are
    /// not, and a worksheet that stores the same position twice reports it
    /// twice, in the order the records appear. Nothing is retained between
    /// calls, so two calls cost two scans.
    ///
    /// # Errors
    ///
    /// Returns the same framing, limit, parse and freshness errors as
    /// [`cell`](Self::cell), or whatever `visitor` returns.
    pub fn visit_cells<F>(&self, visitor: F) -> Result<()>
    where
        F: FnMut(SourceBackedCell) -> Result<()>,
    {
        visit_worksheet_cells(&self.owner, self.sheet_index, None, visitor)
    }

    /// Reports every stored cell of this worksheet from one validated scan,
    /// with cooperative cancellation.
    ///
    /// See [`visit_cells`](Self::visit_cells); this variant checks `execution`
    /// at every point the selected-cell query checks it.
    ///
    /// # Errors
    ///
    /// As [`visit_cells`](Self::visit_cells), plus cancellation.
    pub fn visit_cells_with_execution<F>(
        &self,
        execution: &ExecutionContext,
        visitor: F,
    ) -> Result<()>
    where
        F: FnMut(SourceBackedCell) -> Result<()>,
    {
        visit_worksheet_cells(&self.owner, self.sheet_index, Some(execution), visitor)
    }

    /// Reads one retained worksheet descriptor behind one trailing fence.
    ///
    /// The lookup and `operation` read only retained in-memory state, so they
    /// consume no source bytes and the trailing observation proves what the
    /// leading and trailing pair proved. The missing-worksheet branch fences
    /// before reporting so a changed source still takes precedence over
    /// `WorksheetNotFound`, which is the order the leading fence produced.
    fn metadata<T>(&self, operation: impl FnOnce(&SheetEntry) -> T) -> Result<T> {
        let Some(sheet) = self.owner.sheets.get(self.sheet_index) else {
            self.owner.ensure_current()?;
            return Err(SourceBackedError::WorksheetNotFound(
                self.sheet_index.to_string(),
            ));
        };
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
    // `SharedOleFile::stream_len` resolves a name against the directory tree
    // captured and validated while opening; it consumes no source byte. A
    // leading fence here would therefore prove exactly what the fence on each
    // of the four exits proves, and those exits are where it matters: every
    // one of them observes before it reports, so a changed source still
    // outranks `WorkbookStreamMissing` and the mapped CFB error. This is the
    // shape change 0560 gave the retained-metadata helpers -- fence once,
    // after the value is produced.
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

/// Retained buffer for one workbook-globals scan.
///
/// The buffer holds Workbook stream bytes `[0, filled)` and is grown only by
/// reads that start at `filled`, so a byte is read from the source at most
/// once. Each fill issues one physical read per contiguous FAT run and takes
/// one source-version observation.
///
/// The allocation-chain position is retained across fills in a
/// [`StreamChainHint`]. `read_stream_range` walks the FAT from the stream's
/// **first** sector on every call, so a scan that fills forward in `n` windows
/// walks the prefix `n` times; one hint walks each chain link once for the
/// whole scan. The hint is three integers and a stream identity — it is not an
/// index of the chain — so retaining it allocates nothing and costs no memory
/// proportional to the stream. Everything else about the read, including which
/// physical runs it issues and where it fences the retained source, is
/// unchanged: `read_stream_range` is the same function with a fresh hint.
struct GlobalsBuffer<'a> {
    cfb: &'a SharedOleFile,
    refs: &'a [&'a str],
    stream_len: u64,
    max_global_bytes: u64,
    bytes: Vec<u8>,
    filled: u64,
    /// Allocation-chain position for the next fill. Empty before the first
    /// fill; a hint that does not apply to a fill is ignored by the reader.
    chain: StreamChainHint<'a>,
    /// Size of the next window fill, before clamping.
    window: u64,
    /// Smallest `lbPlyPos` seen in a `BoundSheet8` payload framed *so far*.
    min_sheet_start: Option<u64>,
    /// Cleared permanently once framing contradicts `min_sheet_start`.
    sheet_clamp_active: bool,
}

impl<'a> GlobalsBuffer<'a> {
    fn new(
        cfb: &'a SharedOleFile,
        refs: &'a [&'a str],
        stream_len: u64,
        max_global_bytes: u64,
    ) -> Self {
        Self {
            cfb,
            refs,
            stream_len,
            max_global_bytes,
            bytes: Vec::new(),
            filled: 0,
            chain: cfb.chain_hint(),
            window: GLOBALS_FIRST_WINDOW_BYTES,
            min_sheet_start: None,
            sheet_clamp_active: true,
        }
    }

    /// Upper bound for one window fill. Never below `need`: every check that
    /// gates `need` has already run against header bytes resident in the
    /// buffer, so a fill that stopped short of `need` would not advance and
    /// the scan would not terminate.
    fn fill_cap(&mut self, need: u64) -> u64 {
        if self.sheet_clamp_active
            && self
                .min_sheet_start
                .is_some_and(|sheet_start| need > sheet_start)
        {
            // Globals frame past the smallest declared sheet position, so
            // `lbPlyPos` is corrupt. Drop the clamp for the rest of the scan
            // rather than re-deriving it; `validate_sheet_offsets` still
            // reports the inconsistency with its own error.
            self.sheet_clamp_active = false;
        }
        // `max_global_bytes` bounds retained globals, not the four header
        // bytes that prove a record crosses it: today's order reads that
        // header and then reports "global bytes" with the record's real end.
        let mut cap = self.stream_len.min(self.max_global_bytes.max(need));
        if self.sheet_clamp_active {
            if let Some(sheet_start) = self.min_sheet_start {
                cap = cap.min(sheet_start);
            }
        }
        cap.max(need)
    }

    /// Ensures stream bytes `[0, need)` are resident.
    ///
    /// `exact` reads only the missing bytes; otherwise the fill extends to the
    /// current window, clamped by the stream length, by `max_global_bytes` and
    /// by the smallest known `BoundSheet8` position.
    fn ensure(&mut self, need: u64, exact: bool) -> Result<()> {
        if need <= self.filled {
            return Ok(());
        }
        if need > self.stream_len {
            // Unreachable: `header_end > stream_len` and `end > stream_len`
            // are both rejected before the bytes are requested. Reported as
            // data rather than asserted so a future caller cannot panic here.
            return Err(SourceBackedError::InvalidData(
                "BIFF global fill exceeds the Workbook stream".into(),
            ));
        }
        let end = if exact {
            need
        } else {
            let cap = self.fill_cap(need);
            let end = need.max(self.filled.saturating_add(self.window)).min(cap);
            self.window = self.window.saturating_mul(2).min(GLOBALS_MAX_WINDOW_BYTES);
            end
        };
        let start =
            usize::try_from(self.filled).map_err(|_error| SourceBackedError::ResourceLimit {
                resource: "global address space",
                observed: self.filled,
                maximum: usize::MAX as u64,
            })?;
        let finish = usize::try_from(end).map_err(|_error| SourceBackedError::ResourceLimit {
            resource: "global address space",
            observed: end,
            maximum: usize::MAX as u64,
        })?;
        self.bytes
            .try_reserve_exact(finish - start)
            .map_err(|_error| SourceBackedError::Allocation {
                resource: "workbook globals buffer",
                requested: finish as u64,
            })?;
        self.bytes.resize(finish, 0);
        let cfb: &'a SharedOleFile = self.cfb;
        cfb.read_stream_range_hinted(
            self.refs,
            self.filled,
            &mut self.bytes[start..finish],
            &mut self.chain,
        )
        .map_err(SourceBackedError::from)?;
        self.filled = end;
        Ok(())
    }

    /// Folds one framed `BoundSheet8` stream position into the fill clamp.
    ///
    /// Sheet substreams start at or after the globals end, so no globals byte
    /// lives at or after the smallest position. The clamp is the minimum over
    /// the `BoundSheet8` records framed so far, not over the whole globals:
    /// [MS-XLS] 2.4.28 imposes no ordering on `lbPlyPos`, so a later record may
    /// lower it, and fills issued before that record saw only the higher
    /// bound. A running minimum is still a correct bound on each fill at the
    /// time it is issued, and establishing the true minimum first would need a
    /// second pass over the globals. The value only ever lowers the clamp, so
    /// it cannot invalidate bytes already read; a position at or below `filled`
    /// instead disables the clamp at the next fill, because read bytes cannot
    /// be unread. Malformed payloads raise no error here: the semantic pass
    /// below keeps the existing error order.
    fn note_sheet_start(&mut self, position: u64) {
        self.min_sheet_start = Some(match self.min_sheet_start {
            Some(current) => current.min(position),
            None => position,
        });
    }

    fn header(&self, offset: usize) -> [u8; 4] {
        [
            self.bytes[offset],
            self.bytes[offset + 1],
            self.bytes[offset + 2],
            self.bytes[offset + 3],
        ]
    }
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
    let mut globals = GlobalsBuffer::new(cfb, &refs, stream_len, limits.max_global_bytes);

    // One pass frames the globals and retains the bytes the semantic pass
    // below reads, so no globals byte is read from the source twice.
    //
    // The first `GLOBALS_EXACT_PROLOGUE_RECORDS` records are read exactly.
    // [MS-XLS] 2.1.7.20.1 defines the globals substream as
    // `GLOBALS = BOF [WriteProtect] [FilePass] [Template] ...`, so a FilePass
    // record occupies one of the first three positions and an encrypted
    // workbook is refused with no byte of its payload read. The guarantee
    // actually extends one record further: the last prologue record's fetch
    // buffers the header of record four as well, and `ensure` is a no-op for a
    // header already resident, so a FilePass at index four is also refused
    // before any fill is issued. A FilePass framed later may have payload
    // bytes resident in a fill buffer; they are never framed, interpreted or
    // published, and the refusal is unchanged.
    //
    // From the fifth record on, a fill covers records `i..i+k` before the
    // checks of records `i+1..i+k` run, so a source that fails at a later
    // offset reports its I/O or `SourceChanged` error before a FilePass or
    // limit error of an intermediate record. The prologue couples reads the
    // same way over one record: it fetches record `i`'s payload together with
    // record `i + 1`'s four-byte header, so a source error located in either
    // surfaces during record `i`'s iteration. The per-record check order is
    // unchanged in both phases; only the read coupling is new.
    loop {
        let exact = record_count < GLOBALS_EXACT_PROLOGUE_RECORDS;
        let header_end = offset.checked_add(4).ok_or_else(|| {
            SourceBackedError::InvalidData("BIFF global header offset overflows".into())
        })?;
        if header_end > stream_len {
            return Err(SourceBackedError::InvalidData(
                "truncated BIFF global record header".into(),
            ));
        }
        globals.ensure(header_end, exact)?;
        let header = globals.header(offset as usize);
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
        let Some(next_count) = record_count.checked_add(1) else {
            return Err(SourceBackedError::ResourceLimit {
                resource: "global records",
                observed: u64::MAX,
                maximum: limits.max_global_records as u64,
            });
        };
        record_count = next_count;
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
        // In the exact prologue the payload fetch also covers the next record's
        // four-byte header, so a prologue record costs one read rather than
        // two. The next header is framed at the top of the following
        // iteration, where a FilePass is refused before any ensure for its own
        // payload runs, so no FilePass payload byte is ever requested. The
        // extra four bytes are clamped only by the stream length: today's
        // order also reads a record's header before checking that record
        // against `max_global_bytes`. EOF takes no prefetch, so the prologue
        // never reads past the globals end.
        let need = if exact && kind != EOF {
            end.saturating_add(4).min(stream_len)
        } else {
            end
        };
        globals.ensure(need, exact)?;
        if kind == BOUND_SHEET && payload_len >= 4 {
            let payload = header_end as usize;
            globals.note_sheet_start(u64::from(u32::from_le_bytes([
                globals.bytes[payload],
                globals.bytes[payload + 1],
                globals.bytes[payload + 2],
                globals.bytes[payload + 3],
            ])));
        }
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
    // A window fill may reach past the globals end, never past the stream
    // length, `max_global_bytes` or the smallest known sheet position. Those
    // bytes are dropped here and are never framed or interpreted.
    let mut bytes = globals.bytes;
    bytes.truncate(global_len);

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

struct WorksheetFrame {
    kind: u16,
    payload_len: usize,
}

/// Windowed reader over one worksheet substream.
///
/// The scan keeps stream bytes `[window_start, window_start + window.len())`
/// resident and holds the cursor at the end of that range, so one fill serves
/// the headers and payloads of the records it covers. Every fill is clamped by
/// `fill_bound`, which is the selected sheet's own validated end lowered by
/// `max_worksheet_scan_bytes` measured from the sheet start, and fills start at
/// `window_start + window.len()` and move forward only, so a byte is read from
/// the source at most once. A skipped payload that ends past the filled end is
/// passed with `skip_forward`, which reads nothing.
///
/// The cursor is retained rather than replaced by `read_stream_range`: a
/// worksheet substream starts deep in the Workbook stream, and the cursor keeps
/// its allocation-chain position across fills instead of walking the chain from
/// the first sector on every call.
struct WorksheetScan<'a> {
    cursor: SharedOleStreamCursor<'a>,
    upper_bound: u64,
    /// Highest stream offset any fill may reach.
    fill_bound: u64,
    /// Stream offset of the next byte the scan consumes.
    position: u64,
    /// Stream offset of `window[0]`.
    window_start: u64,
    window: Vec<u8>,
    /// Size of the next window fill, before clamping.
    target: u64,
    scanned_bytes: u64,
    scanned_records: usize,
    limits: SourceBackedLimits,
    execution: Option<&'a ExecutionContext>,
    /// Recycled buffer for payloads a caller keeps across later frames.
    scratch: Vec<u8>,
}

impl<'a> WorksheetScan<'a> {
    /// Opens a scan at `start`, resuming the allocation-chain walk from `chain`.
    ///
    /// `stream_cursor_at` reaches `start` by walking the workbook stream's
    /// allocation chain from its **first** sector, every time: change 0584
    /// priced that at 28,143 links over a 16-sheet workbook's text extraction,
    /// about 91% of them re-walks of a prefix an earlier sheet had already
    /// walked. `chain` is the position change 0585 left open. A document-wide
    /// operation passes one for all of its sheets; a single-sheet operation
    /// passes a fresh one, which is exactly `stream_cursor_at`.
    ///
    /// It is deliberately **not** the shared-string resolver's hint. A hint
    /// retains one position and is discarded when it sits past the offset
    /// asked for, so a single hint alternating between the string table and a
    /// worksheet region would be discarded on every resolve and again on every
    /// sheet, saving nothing on either. The two lifetimes are disjoint for that
    /// reason, and change 0585's record says so.
    fn new(
        cfb: &'a SharedOleFile,
        path: &'a [&'a str],
        start: u64,
        upper_bound: u64,
        limits: SourceBackedLimits,
        execution: Option<&'a ExecutionContext>,
        chain: &mut StreamChainHint<'_>,
    ) -> Result<Self> {
        let cursor = cfb
            .stream_cursor_at_hinted(path, start, chain)
            .map_err(SourceBackedError::from)?;
        Ok(Self {
            cursor,
            upper_bound,
            // `validate_sheet_offsets` runs at open, so `upper_bound` is the
            // selected sheet's validated end before the scan begins and the
            // clamp is exact from the first fill.
            fill_bound: start
                .saturating_add(limits.max_worksheet_scan_bytes)
                .min(upper_bound),
            position: start,
            window_start: start,
            window: Vec::new(),
            target: WORKSHEET_FIRST_WINDOW_BYTES,
            scanned_bytes: 0,
            scanned_records: 0,
            limits,
            execution,
            scratch: Vec::new(),
        })
    }

    /// One past the last stream offset resident in the window, which is also
    /// the cursor's position.
    fn filled_end(&self) -> u64 {
        self.window_start.saturating_add(self.window.len() as u64)
    }

    /// Whether the mean framed bytes per record framed so far exceeds the
    /// density bound.
    ///
    /// A sheet of large records reads ahead more bytes per window than the
    /// window saves in reads, because most of a large record is skipped rather
    /// than framed. The mean is recomputed for every fill, so it is its own
    /// hysteresis.
    fn dense_frames(&self) -> bool {
        self.scanned_records != 0
            && self.scanned_bytes / self.scanned_records as u64 > WORKSHEET_DENSE_FRAME_BYTES
    }

    /// End offset of one fill.
    ///
    /// Never below `need_end`: every check that gates `need_end` has already
    /// run against the sheet boundary and the byte limit, so a fill that
    /// stopped short of `need_end` would not advance and the scan would not
    /// terminate.
    fn fill_end(&mut self, need_end: u64, filled_end: u64) -> u64 {
        if self.dense_frames() {
            self.target = WORKSHEET_FIRST_WINDOW_BYTES;
            return need_end;
        }
        let end = filled_end
            .saturating_add(self.target)
            .clamp(need_end, self.fill_bound.max(need_end));
        self.target = self
            .target
            .saturating_mul(2)
            .min(WORKSHEET_MAX_WINDOW_BYTES);
        end
    }

    /// Ensures stream bytes `[position, need_end)` are resident.
    ///
    /// Every frame and every payload of a sheet whose records sit inside the
    /// current fill takes the resident test and nothing else, so the test is
    /// inlined into the frame loop and only the fill is a call returning a
    /// 48-byte `Result` through memory.
    #[inline]
    fn ensure(&mut self, need_end: u64) -> Result<()> {
        if need_end <= self.filled_end() {
            return Ok(());
        }
        self.fill(need_end)
    }

    /// Reads until stream bytes `[position, need_end)` are resident.
    ///
    /// Bytes already framed are dropped from the front, and one read appends
    /// the fill to the retained tail. The tail is at most one record frame,
    /// because a fill is issued only for bytes the current frame needs.
    #[inline(never)]
    fn fill(&mut self, need_end: u64) -> Result<()> {
        let filled_end = self.filled_end();
        let framed = usize::try_from(self.position.saturating_sub(self.window_start))
            .unwrap_or(usize::MAX)
            .min(self.window.len());
        if framed != 0 {
            self.window.drain(..framed);
            self.window_start = self.window_start.saturating_add(framed as u64);
        }
        let end = self.fill_end(need_end, filled_end);
        let extra = usize::try_from(end.saturating_sub(filled_end)).map_err(|_error| {
            SourceBackedError::ResourceLimit {
                resource: "worksheet window address space",
                observed: end,
                maximum: usize::MAX as u64,
            }
        })?;
        let resident = self.window.len();
        self.window
            .try_reserve_exact(extra)
            .map_err(|_error| SourceBackedError::Allocation {
                resource: "source-backed worksheet window",
                requested: resident.saturating_add(extra) as u64,
            })?;
        self.window.resize(resident + extra, 0);
        if let Err(error) = self.cursor.read_exact(&mut self.window[resident..]) {
            // A failed read leaves the destination undefined, so the window
            // keeps only the bytes earlier fills published.
            self.window.truncate(resident);
            return Err(SourceBackedError::from(error));
        }
        Ok(())
    }

    /// The four header bytes at the scan position, which `ensure` has made
    /// resident.
    ///
    /// One range index in place of four element indexes: `at` is clamped to the
    /// window length, so `at + 4` cannot overflow, and the range panics under
    /// exactly the condition the four element indexes panicked under,
    /// `at + 4 > window.len()`. Only the panic message differs, and `ensure`
    /// makes that state unreachable.
    fn frame_header(&self) -> [u8; 4] {
        let at = usize::try_from(self.position.saturating_sub(self.window_start))
            .unwrap_or(usize::MAX)
            .min(self.window.len());
        let header = &self.window[at..at + 4];
        [header[0], header[1], header[2], header[3]]
    }

    fn next_frame(&mut self) -> Result<WorksheetFrame> {
        self.check_execution()?;
        let cursor_position = self.position;
        if cursor_position
            .checked_add(4)
            .is_none_or(|value| value > self.upper_bound)
        {
            return Err(SourceBackedError::InvalidData(
                "BIFF worksheet has no complete record header before its boundary".into(),
            ));
        }
        // `max_worksheet_scan_bytes` fences the reads as well as the framing:
        // when fewer than four bytes of budget remain the header that would
        // prove the limit is crossed is not read, and the limit is reported
        // against the four bytes it would have taken.
        let header_bytes = self.scanned_bytes.saturating_add(4);
        if header_bytes > self.limits.max_worksheet_scan_bytes {
            return Err(SourceBackedError::ResourceLimit {
                resource: "worksheet scan bytes",
                observed: header_bytes,
                maximum: self.limits.max_worksheet_scan_bytes,
            });
        }
        self.ensure(cursor_position.saturating_add(4))?;
        let header = self.frame_header();
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
        let Some(records) = self.scanned_records.checked_add(1) else {
            return Err(SourceBackedError::ResourceLimit {
                resource: "worksheet scan records",
                observed: u64::MAX,
                maximum: self.limits.max_worksheet_scan_records as u64,
            });
        };
        if records > self.limits.max_worksheet_scan_records {
            return Err(SourceBackedError::ResourceLimit {
                resource: "worksheet scan records",
                observed: records as u64,
                maximum: self.limits.max_worksheet_scan_records as u64,
            });
        }
        let Some(bytes) = self.scanned_bytes.checked_add(frame_len) else {
            return Err(SourceBackedError::ResourceLimit {
                resource: "worksheet scan bytes",
                observed: u64::MAX,
                maximum: self.limits.max_worksheet_scan_bytes,
            });
        };
        if bytes > self.limits.max_worksheet_scan_bytes {
            return Err(SourceBackedError::ResourceLimit {
                resource: "worksheet scan bytes",
                observed: bytes,
                maximum: self.limits.max_worksheet_scan_bytes,
            });
        }
        self.scanned_records = records;
        self.scanned_bytes = bytes;
        self.position = cursor_position.saturating_add(4);
        Ok(WorksheetFrame { kind, payload_len })
    }

    /// Makes one framed payload resident and advances past it, returning where
    /// it sits in the window.
    fn consume_payload(&mut self, frame: &WorksheetFrame) -> Result<std::ops::Range<usize>> {
        self.check_execution()?;
        let end = self.position.saturating_add(frame.payload_len as u64);
        self.ensure(end)?;
        let at = usize::try_from(self.position.saturating_sub(self.window_start))
            .unwrap_or(usize::MAX)
            .min(self.window.len());
        self.position = end;
        Ok(at..at + frame.payload_len)
    }

    fn read_payload(&mut self, frame: &WorksheetFrame) -> Result<&[u8]> {
        let payload = self.consume_payload(frame)?;
        Ok(&self.window[payload])
    }

    fn take_payload(&mut self, frame: &WorksheetFrame) -> Result<Vec<u8>> {
        let payload = self.consume_payload(frame)?;
        let mut taken = std::mem::take(&mut self.scratch);
        taken.clear();
        taken
            .try_reserve_exact(payload.len())
            .map_err(|_error| SourceBackedError::Allocation {
                resource: "source-backed worksheet payload",
                requested: payload.len() as u64,
            })?;
        taken.extend_from_slice(&self.window[payload]);
        Ok(taken)
    }

    fn recycle_payload(&mut self, payload: Vec<u8>) {
        self.scratch = payload;
    }

    /// Advances past a payload the caller does not frame, interpret or
    /// publish.
    ///
    /// No read is issued for it. Bytes a fill taken to frame earlier records
    /// already carries are passed in the window; a payload that ends past the
    /// filled end is passed by moving the cursor, which reads nothing, and the
    /// window is dropped so the next fill starts after the payload.
    fn skip_payload(&mut self, frame: &WorksheetFrame) -> Result<()> {
        self.check_execution()?;
        let end = self.position.saturating_add(frame.payload_len as u64);
        let filled_end = self.filled_end();
        if end <= filled_end {
            self.position = end;
            return Ok(());
        }
        self.cursor
            .skip_forward(end.saturating_sub(filled_end))
            .map_err(SourceBackedError::from)?;
        self.window.clear();
        self.window_start = end;
        self.position = end;
        Ok(())
    }

    /// Honours the caller's cooperative cancellation, twice per framed record.
    ///
    /// Most callers pass no execution context, and that case is a null test
    /// which is inlined into the frame loop rather than a call returning a
    /// 48-byte `Result` through memory.
    #[inline]
    fn check_execution(&self) -> Result<()> {
        match self.execution {
            None => Ok(()),
            Some(context) => context.check().map_err(SourceBackedError::from),
        }
    }
}

fn check_text_state(owner: &SourceInner, execution: Option<&ExecutionContext>) -> Result<()> {
    check_text_cancellation(execution)?;
    owner.ensure_current()
}

/// The cancellation half of [`check_text_state`], for the row loop, where the
/// source observation is taken one call deeper.
///
/// [`write_text_sheet`] emits rows out of the map [`scan_text_sheet`] already
/// collected and fenced; it consumes no source byte. Every byte it emits goes
/// through [`SourceCheckedTextSink`], which observes the source immediately
/// before the write and again immediately after it, so a per-row observation
/// here would sit between two in-memory steps and re-prove what the sink's
/// leading observation proves a moment later. Cancellation is not in the same
/// position: a row whose object the writer skips reaches no sink call at all,
/// so the row loop keeps its own cancellation granularity.
fn check_text_cancellation(execution: Option<&ExecutionContext>) -> Result<()> {
    if let Some(context) = execution {
        context.check().map_err(SourceBackedError::from)?;
    }
    Ok(())
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
    strings: &mut SharedStringResolver<'_>,
    sheet_chain: &mut StreamChainHint<'_>,
) -> Result<SourceTextSheet> {
    let mut collected = SourceTextSheet::new();
    let mut scan = WorksheetScan::new(
        &owner.cfb,
        refs,
        sheet.start,
        sheet.end,
        owner.limits,
        execution,
        sheet_chain,
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
                collect_source_cell(&formula, owner, &mut collected, execution, strings)?;
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
                    collect_source_cell(&cell, owner, &mut collected, execution, strings)?;
                }
            },
            0x0201 | 0x0203 | 0x0204 | 0x0205 | 0x027E | 0x00FD => {
                let payload = scan.read_payload(&frame)?;
                let cell = CellRecord::parse(frame.kind, payload, &owner.encoding)
                    .map_err(SourceBackedError::Parse)?;
                collect_source_cell(&cell, owner, &mut collected, execution, strings)?;
            },
            0x00BD => {
                let payload = scan.read_payload(&frame)?;
                let mut processing = Ok(());
                CellRecord::visit_mul_rk(payload, |cell| {
                    if processing.is_ok() {
                        processing =
                            collect_source_cell(&cell, owner, &mut collected, execution, strings);
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
                        processing =
                            collect_source_cell(&cell, owner, &mut collected, execution, strings);
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
    sheet: &SourceTextSheet,
    writer: &mut SequentialTextWriter<'options, 'output, W>,
    execution: Option<&ExecutionContext>,
) -> std::result::Result<(), TextOutputError<SourceBackedError>> {
    let mut row = 0_u16;
    loop {
        check_text_cancellation(execution).map_err(|source| writer.document_error(source))?;
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

/// What one worksheet scan does with each cell record it parses.
///
/// A selected-cell query and a whole-sheet walk run the *identical* frame
/// loop over the identical checks in the identical order; they differ only in
/// what they do with a parsed record, and in whether a string-valued `FORMULA`
/// has to be held back until its `STRING` result arrives. Monomorphizing the
/// loop over this trait keeps the selected-cell path exactly as it was while
/// giving the whole-sheet walk **one** validated scan instead of one per cell.
trait CellSink {
    /// Consumes one parsed cell record.
    ///
    /// Every implementation validates the record's XF index first, for every
    /// record, so that a malformed cell format is refused wherever it sits.
    fn accept(
        &mut self,
        record: &CellRecord,
        scan: &ScanContext<'_>,
        strings: &mut SharedStringResolver<'_>,
    ) -> Result<()>;

    /// Whether this string-valued `FORMULA` must be held until its `STRING`
    /// result arrives.
    ///
    /// A selected-cell query holds back only the target, because every other
    /// cell's value is discarded anyway; a whole-sheet walk holds back all of
    /// them, because it reports every one.
    fn defers_string_formula(&self, record: &CellRecord) -> bool;

    /// Whether the sink reads anything out of the record at this position
    /// beyond its XF index.
    ///
    /// A sink that answers `false` still gets the record **validated** — the
    /// scan hands the payload to [`CellRecord::measure`], which runs every
    /// check `CellRecord::parse` runs — and then gets
    /// [`CellSink::accept_measured`] instead of `accept`. Nothing is skipped;
    /// what is skipped is the `String` a `Label` transcodes, the `Vec<u8>` a
    /// `Formula` copies, and the 88-byte record that carries them.
    ///
    /// The scan decides this from the four header bytes every cell record
    /// opens with, which both parses read identically, so the answer can never
    /// change which bytes are checked.
    fn wants(&self, row: u16, col: u16) -> bool;

    /// Consumes one validated cell record the sink does not want.
    ///
    /// Every implementation validates the XF index, exactly as
    /// [`CellSink::accept`] does and in the same position, so that a malformed
    /// cell format is refused wherever it sits whether or not the sink keeps
    /// the cell.
    fn accept_measured(&mut self, measured: &MeasuredCell, scan: &ScanContext<'_>) -> Result<()> {
        scan.formatting
            .validate_cell_xf(measured.xf_index)
            .map_err(SourceBackedError::Parse)
    }
}

/// The per-scan constants every sink needs, gathered once so that the frame
/// loop hands a sink two pointers instead of four.
struct ScanContext<'a> {
    owner: &'a SourceInner,
    formatting: &'a Formatting,
    execution: Option<&'a ExecutionContext>,
}

/// The selected-cell sink: exactly what `process_cell` did before this scan
/// loop was shared.
struct TargetCell {
    row: u16,
    column: u16,
    found: Option<SourceBackedCell>,
}

impl CellSink for TargetCell {
    #[inline]
    fn accept(
        &mut self,
        record: &CellRecord,
        scan: &ScanContext<'_>,
        strings: &mut SharedStringResolver<'_>,
    ) -> Result<()> {
        scan.formatting
            .validate_cell_xf(cell_xf_index(record))
            .map_err(SourceBackedError::Parse)?;
        if record.row() != self.row || record.col() != self.column {
            return Ok(());
        }
        let Some(cell) =
            Cell::from_record_with_formula_context(record, None, None, Some(scan.formatting))
        else {
            return Ok(());
        };
        let value = if let Some(string_index) = cell.shared_string_index() {
            resolve_shared_string(scan.owner, string_index, scan.execution, strings)?
        } else {
            cell.value().clone()
        };
        self.found = Some(SourceBackedCell {
            row: u32::from(self.row),
            column: u32::from(self.column),
            value,
        });
        Ok(())
    }

    #[inline]
    fn defers_string_formula(&self, record: &CellRecord) -> bool {
        record.row() == self.row && record.col() == self.column
    }

    #[inline]
    fn wants(&self, row: u16, col: u16) -> bool {
        row == self.row && col == self.column
    }
}

/// The whole-sheet sink: reports every stored cell in stream order and retains
/// nothing of its own.
struct VisitCells<F> {
    visitor: F,
}

impl<F> CellSink for VisitCells<F>
where
    F: FnMut(SourceBackedCell) -> Result<()>,
{
    fn accept(
        &mut self,
        record: &CellRecord,
        scan: &ScanContext<'_>,
        strings: &mut SharedStringResolver<'_>,
    ) -> Result<()> {
        scan.formatting
            .validate_cell_xf(cell_xf_index(record))
            .map_err(SourceBackedError::Parse)?;
        let Some(cell) =
            Cell::from_record_with_formula_context(record, None, None, Some(scan.formatting))
        else {
            return Ok(());
        };
        let value = if let Some(string_index) = cell.shared_string_index() {
            resolve_shared_string(scan.owner, string_index, scan.execution, strings)?
        } else {
            cell.value().clone()
        };
        (self.visitor)(SourceBackedCell {
            row: u32::from(record.row()),
            column: u32::from(record.col()),
            value,
        })
    }

    fn defers_string_formula(&self, _record: &CellRecord) -> bool {
        true
    }

    /// A whole-sheet walk reports every stored cell, so it wants every record
    /// materialized and never reaches the measure-only path.
    fn wants(&self, _row: u16, _col: u16) -> bool {
        true
    }
}

/// Whether `sink` wants the cell record in `payload` materialized.
///
/// Every BIFF8 cell record this scan parses opens with `rw` and `col`, two
/// little-endian `u16`s at offsets 0 and 2 (`[MS-XLS]` 2.5.19 `Cell`), so the
/// position is readable before either parse commits. A payload too short to
/// carry them is answered `true`, which hands it to the materializing parse and
/// therefore to exactly the refusal it always produced; the measure-only path
/// never sees a record whose length checks have not been decided by the parse
/// they belong to.
#[inline]
fn wants_record<S: CellSink + ?Sized>(sink: &S, payload: &[u8]) -> bool {
    payload.get(..4).is_none_or(|head| {
        sink.wants(
            u16::from_le_bytes([head[0], head[1]]),
            u16::from_le_bytes([head[2], head[3]]),
        )
    })
}

/// Scans one worksheet substream to its EOF, handing every parsed cell record
/// to `sink`.
///
/// The caller has already taken the leading cancellation check and freshness
/// fence; this function takes the trailing pair when the scan reaches EOF, so
/// that every operation built on it keeps the same two-observation bracket a
/// selected-cell query always had.
fn scan_worksheet<S: CellSink>(
    owner: &Arc<SourceInner>,
    sheet_index: usize,
    execution: Option<&ExecutionContext>,
    sink: &mut S,
) -> Result<()> {
    let sheet = owner
        .sheets
        .get(sheet_index)
        .ok_or_else(|| SourceBackedError::WorksheetNotFound(sheet_index.to_string()))?;
    let refs = owner
        .workbook_path
        .iter()
        .map(String::as_str)
        .collect::<Vec<_>>();
    let mut strings = SharedStringResolver::new(owner, &refs);
    let mut pending_formula = None;
    let context = ScanContext {
        owner,
        formatting: &owner.formatting,
        execution,
    };
    // A selected-cell query and a whole-sheet walk each construct one cursor,
    // so there is no earlier position on this stream to resume from: a fresh
    // hint is exactly the cold walk `stream_cursor_at` performed.
    let mut sheet_chain = owner.cfb.chain_hint();
    let mut scan = WorksheetScan::new(
        &owner.cfb,
        &refs,
        sheet.start,
        sheet.end,
        owner.limits,
        execution,
        &mut sheet_chain,
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
            return finish_scan(owner, execution);
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
                sink.accept(&formula, &context, &mut strings)?;
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
                if !wants_record(sink, payload) {
                    let measured = CellRecord::measure(frame.kind, payload, &owner.encoding)
                        .map_err(SourceBackedError::Parse)?;
                    sink.accept_measured(&measured, &context)?;
                    continue;
                }
                let cell = CellRecord::parse(frame.kind, payload, &owner.encoding)
                    .map_err(SourceBackedError::Parse)?;
                if matches!(
                    cell,
                    CellRecord::Formula {
                        value: FormulaValue::StringPending,
                        ..
                    }
                ) {
                    if sink.defers_string_formula(&cell) {
                        pending_formula = Some(cell);
                    } else {
                        sink.accept(&cell, &context, &mut strings)?;
                    }
                } else {
                    sink.accept(&cell, &context, &mut strings)?;
                }
            },
            0x0201 | 0x0203 | 0x0204 | 0x0205 | 0x027E | 0x00FD => {
                let payload = scan.read_payload(&frame)?;
                if !wants_record(sink, payload) {
                    let measured = CellRecord::measure(frame.kind, payload, &owner.encoding)
                        .map_err(SourceBackedError::Parse)?;
                    sink.accept_measured(&measured, &context)?;
                    continue;
                }
                let cell = CellRecord::parse(frame.kind, payload, &owner.encoding)
                    .map_err(SourceBackedError::Parse)?;
                sink.accept(&cell, &context, &mut strings)?;
            },
            0x00BD => {
                let payload = scan.read_payload(&frame)?;
                let mut processing = Ok(());
                CellRecord::visit_mul_rk(payload, |cell| {
                    if processing.is_ok() {
                        processing = sink.accept(&cell, &context, &mut strings);
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
                        processing = sink.accept(&cell, &context, &mut strings);
                    }
                })
                .map_err(SourceBackedError::Parse)?;
                processing?;
            },
            _ => scan.skip_payload(&frame)?,
        }
    }
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
    let mut sink = TargetCell {
        row: row as u16,
        column: column as u16,
        found: None,
    };
    scan_worksheet(owner, sheet_index, execution, &mut sink)?;
    Ok(sink.found)
}

/// Walks one worksheet once and reports every stored cell in stream order.
fn visit_worksheet_cells<F>(
    owner: &Arc<SourceInner>,
    sheet_index: usize,
    execution: Option<&ExecutionContext>,
    visitor: F,
) -> Result<()>
where
    F: FnMut(SourceBackedCell) -> Result<()>,
{
    if let Some(context) = execution {
        context.check().map_err(SourceBackedError::from)?;
    }
    owner.ensure_current()?;
    let mut sink = VisitCells { visitor };
    scan_worksheet(owner, sheet_index, execution, &mut sink)
}

fn finish_scan(owner: &SourceInner, execution: Option<&ExecutionContext>) -> Result<()> {
    if let Some(context) = execution {
        context.check().map_err(SourceBackedError::from)?;
    }
    owner.ensure_current()
}

/// The per-scan state `resolve_shared_string` would otherwise rebuild on every
/// string cell.
///
/// One sheet scan resolves one shared string per `LabelSst` -- 658 times on
/// `ConditionalFormattingSamples.xls`, 16,055 times on `54016.xls` -- and each
/// of those calls was rebuilding the workbook stream path into a fresh `Vec`
/// and walking the `Workbook` allocation chain from its **first** sector to
/// reach the string table. Both are per-scan constants, so both live here.
///
/// The chain position is dedicated to the SST region and is deliberately
/// **not** shared with the worksheet scan's own cursor. A hint retains one
/// position, and the worksheet cursor sits permanently past the string table:
/// one hint serving both would be discarded as a backward step by every
/// resolve and again by every sheet, and would save nothing on either path.
struct SharedStringResolver<'a> {
    /// The workbook stream path, borrowed for the whole scan.
    path: &'a [&'a str],
    /// Allocation-chain position of the last resolved entry. Empty before the
    /// first resolve; a hint that does not apply is ignored by the reader,
    /// which then walks from the stream's first sector exactly as before.
    chain: StreamChainHint<'a>,
}

impl<'a> SharedStringResolver<'a> {
    fn new(owner: &'a SourceInner, path: &'a [&'a str]) -> Self {
        Self {
            path,
            chain: owner.cfb.chain_hint(),
        }
    }
}

/// Resolves one shared string, refusing a mutation over any other failure.
///
/// Every byte this returns is read through `SharedOleStreamCursor::read_exact`,
/// which observes the source after each read (change 0558), so the value is
/// already bracketed by the observation the scan took before it and by that
/// trailing observation. The resolver therefore takes no observation of its own
/// on the path that succeeds. It keeps one on the path that fails, where the
/// observation is not redundant but decisive: change 0317's precedence makes a
/// stale source outrank the locator, allocation, chain and decode errors a
/// mutation can provoke. As in `SharedOleFile::finish_stream_range`, only a
/// `SourceChanged` refusal displaces the original error; an observation that
/// cannot be taken at all leaves the original error in place.
fn resolve_shared_string(
    owner: &SourceInner,
    string_index: u32,
    execution: Option<&ExecutionContext>,
    strings: &mut SharedStringResolver<'_>,
) -> Result<litchi_core::sheet::CellValue> {
    match resolve_shared_string_inner(owner, string_index, execution, strings) {
        Ok(value) => Ok(value),
        Err(original) => match owner.ensure_current() {
            Err(changed @ SourceBackedError::SourceChanged { .. }) => Err(changed),
            _ => Err(original),
        },
    }
}

fn resolve_shared_string_inner(
    owner: &SourceInner,
    string_index: u32,
    execution: Option<&ExecutionContext>,
    strings: &mut SharedStringResolver<'_>,
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
    let mut cursor = owner
        .cfb
        .stream_cursor_at_hinted(strings.path, first_offset, &mut strings.chain)
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
        Ok(value) => Ok(litchi_core::sheet::CellValue::String(value)),
        Err(error) => Err(map_shared_string_error(error)),
    }
}

fn decode_source_cell(
    record: &CellRecord,
    owner: &SourceInner,
    formatting: &Formatting,
    execution: Option<&ExecutionContext>,
    strings: &mut SharedStringResolver<'_>,
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
            resolve_shared_string(owner, string_index, execution, strings)?
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
    strings: &mut SharedStringResolver<'_>,
) -> Result<()> {
    if let Some((row, column, value)) =
        decode_source_cell(record, owner, &owner.formatting, execution, strings)?
    {
        collected.insert(row, column, value, owner.limits)?;
    }
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
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic by design"
    )]

    use super::*;
    use std::io;
    use std::sync::Mutex;

    /// A positional source that records every read range and every
    /// source-version observation, so two legs can be compared for I/O
    /// identity rather than only for the bytes they produce.
    #[derive(Debug)]
    struct RecordingSource {
        bytes: Vec<u8>,
        ranges: Mutex<Vec<(u64, usize)>>,
        versions: Mutex<usize>,
    }

    impl RecordingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                ranges: Mutex::new(Vec::new()),
                versions: Mutex::new(0),
            }
        }

        fn take(&self) -> (Vec<(u64, usize)>, usize) {
            let ranges = std::mem::take(&mut *self.ranges.lock().unwrap());
            let versions = std::mem::replace(&mut *self.versions.lock().unwrap(), 0);
            (ranges, versions)
        }
    }

    impl ReadAt for RecordingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let Ok(start) = usize::try_from(offset) else {
                return Ok(0);
            };
            if start >= self.bytes.len() || output.is_empty() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            self.ranges.lock().unwrap().push((offset, count));
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            *self.versions.lock().unwrap() += 1;
            Ok(SourceVersion::new(0x584c_535f_4755_524d, 0))
        }
    }

    fn erased(source: &Arc<RecordingSource>) -> Arc<dyn ReadAt> {
        let source: Arc<RecordingSource> = Arc::clone(source);
        source
    }

    fn ole_fixture(name: &str) -> Vec<u8> {
        std::fs::read(
            std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ole/xls")
                .join(name),
        )
        .unwrap()
    }

    /// The fill schedule below is issued twice over the same file: once
    /// through `GlobalsBuffer`, which carries one chain hint across all of its
    /// fills, and once through plain `SharedOleFile::read_stream_range`, which
    /// walks the chain from the stream's first sector on every call. The two
    /// legs must be indistinguishable from outside the reader.
    const FILL_SCHEDULE: [(u64, bool); 7] = [
        (4, true),
        (24, true),
        (60, true),
        (512, false),
        (4_096, false),
        (12_000, false),
        (30_000, false),
    ];

    #[test]
    fn the_globals_scan_reads_identically_with_and_without_its_chain_hint() {
        let bytes = ole_fixture("SimpleWithImages-mac.xls");
        let source = Arc::new(RecordingSource::new(bytes));
        let cfb = SharedOleFile::open(erased(&source)).unwrap();
        let stream_len = cfb.stream_len(&["Workbook"]).unwrap();
        assert!(
            stream_len > FILL_SCHEDULE.last().unwrap().0,
            "the fixture must be long enough to need every fill"
        );
        let refs = ["Workbook"];

        let mut globals = GlobalsBuffer::new(&cfb, &refs, stream_len, u64::MAX);
        let _ = source.take();
        for (need, exact) in FILL_SCHEDULE {
            globals.ensure(need, exact).unwrap();
        }
        let (hinted_ranges, hinted_versions) = source.take();
        let hinted_bytes = globals.bytes.clone();
        let filled = globals.filled;

        let plain_source = Arc::new(RecordingSource::new(source.bytes.clone()));
        let plain_cfb = SharedOleFile::open(erased(&plain_source)).unwrap();
        let mut plain_bytes = vec![0u8; usize::try_from(filled).unwrap()];
        let _ = plain_source.take();
        for (start, end) in fill_extents(&globals_extents(stream_len)) {
            let (start, end) = (
                usize::try_from(start).unwrap(),
                usize::try_from(end).unwrap(),
            );
            plain_cfb
                .read_stream_range(&refs, start as u64, &mut plain_bytes[start..end])
                .unwrap();
        }
        let (plain_ranges, plain_versions) = plain_source.take();

        assert_eq!(hinted_bytes, plain_bytes, "payload bytes differ");
        assert_eq!(hinted_ranges, plain_ranges, "positional reads differ");
        assert_eq!(
            hinted_versions, plain_versions,
            "source-version observations differ"
        );
        assert!(
            hinted_ranges.len() >= FILL_SCHEDULE.len(),
            "the schedule must actually reach the source"
        );
    }

    /// Recomputes the `[filled, end)` extents `GlobalsBuffer` produces for
    /// `FILL_SCHEDULE`, independently of the buffer itself, so the comparison
    /// leg shares no code with the leg under test.
    fn globals_extents(stream_len: u64) -> Vec<u64> {
        let mut ends = Vec::new();
        let mut filled = 0u64;
        let mut window = GLOBALS_FIRST_WINDOW_BYTES;
        for (need, exact) in FILL_SCHEDULE {
            if need <= filled {
                continue;
            }
            let end = if exact {
                need
            } else {
                let end = need
                    .max(filled.saturating_add(window))
                    .min(stream_len.max(need));
                window = window.saturating_mul(2).min(GLOBALS_MAX_WINDOW_BYTES);
                end
            };
            ends.push(end);
            filled = end;
        }
        ends
    }

    fn fill_extents(ends: &[u64]) -> Vec<(u64, u64)> {
        let mut extents = Vec::new();
        let mut start = 0u64;
        for &end in ends {
            extents.push((start, end));
            start = end;
        }
        extents
    }

    #[test]
    fn limits_are_finite_by_default() {
        let limits = SourceBackedLimits::default();
        assert!(limits.max_input_bytes > 0);
        assert!(limits.max_global_bytes > 0);
        assert!(limits.max_worksheet_scan_bytes > 0);
    }
}
