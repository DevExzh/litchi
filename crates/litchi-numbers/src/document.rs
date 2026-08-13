//! Archive-free Numbers document semantics.

use std::collections::HashSet;
use std::fmt;
use std::fmt::Write as _;
use std::path::Path;
use std::sync::Arc;

use thiserror::Error as ThisError;

use crate::cell::Value;
use crate::{Sheet, SheetSelector};

/// Maximum number of ordered sheets retained by one semantic document.
pub const MAX_SHEETS: usize = 4096;
/// Maximum number of tables retained by one semantic document.
pub const MAX_TABLES: usize = 65_536;
/// Maximum number of materialized cells retained by one semantic document.
pub const MAX_MATERIALIZED_CELLS: usize = 16_000_000;
/// Maximum UTF-8 bytes retained by semantic names, headers, and textual cells.
pub const DEFAULT_MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// Stable source-capture resource category for archive-free Numbers ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DocumentSourceLimitKind {
    /// Encoded bytes supplied by the caller.
    InputBytes,
    /// Independently addressed items discovered in the source.
    Entries,
    /// Decoded bytes contributed by one source item.
    EntryBytes,
    /// Aggregate decoded bytes contributed by all source items.
    AggregateBytes,
    /// Bytes contributed by one document payload component.
    ComponentBytes,
}

impl fmt::Display for DocumentSourceLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::AggregateBytes => "aggregate bytes",
            Self::ComponentBytes => "component bytes",
        })
    }
}

/// Invalid caller-selected source-capture limits.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[error(
    "Numbers document source {kind} limit must be non-zero and no greater than {maximum}, got {value}"
)]
#[non_exhaustive]
pub struct DocumentSourceLimitsError {
    /// Resource category whose requested limit is invalid.
    pub kind: DocumentSourceLimitKind,
    /// Requested resource ceiling.
    pub value: u64,
    /// Format-wide hard ceiling for this resource.
    pub maximum: u64,
}

/// Checked physical resource ceilings for archive-free Numbers source capture.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocumentSourceLimits {
    input_bytes: u64,
    entries: usize,
    entry_bytes: u64,
    aggregate_bytes: u64,
    component_bytes: usize,
}

impl DocumentSourceLimits {
    /// Build a checked source-capture resource profile.
    ///
    /// # Errors
    ///
    /// Returns an error if any ceiling is zero or exceeds its hard maximum.
    pub fn new(
        max_input_bytes: u64,
        max_entries: usize,
        max_entry_bytes: u64,
        max_aggregate_bytes: u64,
        max_component_bytes: usize,
    ) -> std::result::Result<Self, DocumentSourceLimitsError> {
        check_source_limit(
            DocumentSourceLimitKind::InputBytes,
            max_input_bytes,
            litchi_iwa_detect::Limits::HARD_MAX_INPUT_BYTES,
        )?;
        check_source_limit(
            DocumentSourceLimitKind::Entries,
            usize_u64(max_entries),
            usize_u64(litchi_iwa_detect::Limits::HARD_MAX_FILES),
        )?;
        check_source_limit(
            DocumentSourceLimitKind::EntryBytes,
            max_entry_bytes,
            litchi_iwa_detect::Limits::HARD_MAX_ENTRY_SIZE,
        )?;
        check_source_limit(
            DocumentSourceLimitKind::AggregateBytes,
            max_aggregate_bytes,
            litchi_iwa_detect::Limits::HARD_MAX_TOTAL_SIZE,
        )?;
        check_source_limit(
            DocumentSourceLimitKind::ComponentBytes,
            usize_u64(max_component_bytes),
            usize_u64(litchi_iwa_detect::Limits::HARD_MAX_IWA_STREAM_SIZE),
        )?;
        Ok(Self {
            input_bytes: max_input_bytes,
            entries: max_entries,
            entry_bytes: max_entry_bytes,
            aggregate_bytes: max_aggregate_bytes,
            component_bytes: max_component_bytes,
        })
    }

    /// Maximum encoded bytes accepted from one source.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.input_bytes
    }

    /// Maximum independently addressed source items accepted.
    #[must_use]
    pub const fn max_entries(self) -> usize {
        self.entries
    }

    /// Maximum decoded bytes accepted from one source item.
    #[must_use]
    pub const fn max_entry_bytes(self) -> u64 {
        self.entry_bytes
    }

    /// Maximum aggregate decoded bytes accepted from all source items.
    #[must_use]
    pub const fn max_aggregate_bytes(self) -> u64 {
        self.aggregate_bytes
    }

    /// Maximum bytes accepted from one document payload component.
    #[must_use]
    pub const fn max_component_bytes(self) -> usize {
        self.component_bytes
    }

    fn detector_limits(self) -> litchi_iwa_detect::Result<litchi_iwa_detect::Limits> {
        litchi_iwa_detect::Limits::new(
            self.input_bytes,
            self.entries,
            self.entry_bytes,
            self.aggregate_bytes,
            self.component_bytes,
        )
    }
}

impl Default for DocumentSourceLimits {
    fn default() -> Self {
        let limits = litchi_iwa_detect::Limits::default();
        Self {
            input_bytes: limits.max_input_bytes(),
            entries: limits.max_files(),
            entry_bytes: limits.max_entry_size(),
            aggregate_bytes: limits.max_total_size(),
            component_bytes: limits.max_iwa_stream_size(),
        }
    }
}

/// Physical and semantic profiles for archive-free Numbers document ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DocumentReadOptions {
    source: DocumentSourceLimits,
    semantic: Limits,
}

impl DocumentReadOptions {
    /// Combine checked source-capture and semantic resource profiles.
    #[must_use]
    pub const fn new(source: DocumentSourceLimits, semantic: Limits) -> Self {
        Self { source, semantic }
    }

    /// Return the bounded source-capture profile.
    #[must_use]
    pub const fn source(self) -> DocumentSourceLimits {
        self.source
    }

    /// Return the semantic projection profile.
    #[must_use]
    pub const fn semantic(self) -> Limits {
        self.semantic
    }
}

/// Content-free filesystem failure category for archive-free ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum IoKind {
    /// The source does not exist.
    NotFound,
    /// Access to the source was denied.
    PermissionDenied,
    /// The operation conflicted with an existing resource.
    AlreadyExists,
    /// The caller supplied an invalid input.
    InvalidInput,
    /// The source contained invalid data.
    InvalidData,
    /// The operation timed out.
    TimedOut,
    /// The operation was interrupted.
    Interrupted,
    /// The source ended unexpectedly.
    UnexpectedEof,
    /// Another content-free I/O category.
    Other,
}

impl fmt::Display for IoKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::NotFound => "not found",
            Self::PermissionDenied => "permission denied",
            Self::AlreadyExists => "already exists",
            Self::InvalidInput => "invalid input",
            Self::InvalidData => "invalid data",
            Self::TimedOut => "timed out",
            Self::Interrupted => "interrupted",
            Self::UnexpectedEof => "unexpected end of input",
            Self::Other => "other I/O failure",
        })
    }
}

/// Content-free resource category reported by archive-free ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ReadLimitKind {
    /// Complete captured input bytes.
    InputBytes,
    /// Independently addressed source items.
    Entries,
    /// Encoded or decoded semantic payload bytes.
    PayloadBytes,
    /// Fields or field-like records in semantic payloads.
    PayloadFields,
    /// Nested semantic payload traversal depth.
    PayloadNesting,
    /// Aggregate semantic traversal or rewrite work.
    PayloadWork,
    /// Rooted semantic sheets.
    Sheets,
    /// Rooted semantic tables.
    Tables,
    /// Materialized semantic cells.
    Cells,
    /// Aggregate semantic UTF-8 text bytes.
    TextBytes,
    /// A future content-free resource category not known by this release.
    Other,
}

impl fmt::Display for ReadLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::Entries => "entries",
            Self::PayloadBytes => "payload bytes",
            Self::PayloadFields => "payload fields",
            Self::PayloadNesting => "payload nesting",
            Self::PayloadWork => "payload work",
            Self::Sheets => "sheets",
            Self::Tables => "tables",
            Self::Cells => "cells",
            Self::TextBytes => "text bytes",
            Self::Other => "other resource",
        })
    }
}

/// Errors raised while publishing an archive-free Numbers document.
///
/// Display, Debug, and error-source chains contain only closed categories and
/// numeric bounds. Source paths, authored content, and lower-layer diagnostic
/// strings are deliberately discarded.
#[derive(Debug, ThisError, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum ReadError {
    /// A filesystem operation failed.
    #[error("Numbers document I/O failure: {kind}")]
    Io { kind: IoKind },
    /// The recognized source belongs to a different application.
    #[error("iWork source is not a Numbers document")]
    NotNumbers,
    /// A checked resource ceiling was exceeded.
    #[error("Numbers document {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: ReadLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// The source could not be captured safely.
    #[error("invalid Numbers document source")]
    InvalidSource,
    /// The captured source is not a valid Numbers document.
    #[error("invalid Numbers document format")]
    InvalidFormat,
    /// A bounded allocation failed.
    #[error("Numbers document allocation failed for {amount} units")]
    Allocation {
        /// Requested elements or bytes.
        amount: usize,
    },
}

/// Deterministic measurements retained for a source-backed document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Stats {
    /// Number of source records validated during projection.
    pub source_record_count: usize,
    /// Number of rooted semantic sheets.
    pub sheet_count: usize,
    /// Number of rooted semantic tables.
    pub table_count: usize,
}

#[derive(Debug)]
struct SourceDiagnostics {
    metadata: litchi_core::Metadata,
    stats: Stats,
}

/// Caller-selected semantic resource governed by [`Limits`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Rooted sheets.
    Sheets,
    /// Rooted tables.
    Tables,
    /// Materialized cells.
    MaterializedCells,
    /// Retained and rendered UTF-8 bytes.
    TextBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Sheets => "sheets",
            Self::Tables => "tables",
            Self::MaterializedCells => "materialized cells",
            Self::TextBytes => "text bytes",
        })
    }
}

/// An invalid semantic document ceiling.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[error("Numbers document {kind} limit must be no greater than {maximum}, got {value}")]
#[non_exhaustive]
pub struct LimitsError {
    /// Resource category whose requested ceiling is invalid.
    pub kind: LimitKind,
    /// Requested ceiling.
    pub value: usize,
    /// Hard semantic ceiling.
    pub maximum: usize,
}

/// Finite construction limits for an immutable Numbers document.
#[allow(
    clippy::struct_field_names,
    reason = "The public budget accessors intentionally share one max_* vocabulary"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_sheets: usize,
    max_tables: usize,
    max_materialized_cells: usize,
    max_text_bytes: usize,
}

impl Limits {
    /// Creates a checked profile. Zero is a valid exact tightening.
    ///
    /// # Errors
    ///
    /// Returns an error when a requested ceiling exceeds its hard maximum.
    pub const fn new(
        max_sheets: usize,
        max_tables: usize,
        max_materialized_cells: usize,
        max_text_bytes: usize,
    ) -> std::result::Result<Self, LimitsError> {
        if max_sheets > MAX_SHEETS {
            return Err(LimitsError {
                kind: LimitKind::Sheets,
                value: max_sheets,
                maximum: MAX_SHEETS,
            });
        }
        if max_tables > MAX_TABLES {
            return Err(LimitsError {
                kind: LimitKind::Tables,
                value: max_tables,
                maximum: MAX_TABLES,
            });
        }
        if max_materialized_cells > MAX_MATERIALIZED_CELLS {
            return Err(LimitsError {
                kind: LimitKind::MaterializedCells,
                value: max_materialized_cells,
                maximum: MAX_MATERIALIZED_CELLS,
            });
        }
        if max_text_bytes > DEFAULT_MAX_TEXT_BYTES {
            return Err(LimitsError {
                kind: LimitKind::TextBytes,
                value: max_text_bytes,
                maximum: DEFAULT_MAX_TEXT_BYTES,
            });
        }
        Ok(Self {
            max_sheets,
            max_tables,
            max_materialized_cells,
            max_text_bytes,
        })
    }

    /// Returns the configured sheet ceiling.
    #[must_use]
    pub const fn max_sheets(self) -> usize {
        self.max_sheets
    }

    /// Returns the configured table ceiling.
    #[must_use]
    pub const fn max_tables(self) -> usize {
        self.max_tables
    }

    /// Returns the configured materialized-cell ceiling.
    #[must_use]
    pub const fn max_materialized_cells(self) -> usize {
        self.max_materialized_cells
    }

    /// Returns the configured text-byte ceiling.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_sheets: MAX_SHEETS,
            max_tables: MAX_TABLES,
            max_materialized_cells: MAX_MATERIALIZED_CELLS,
            max_text_bytes: DEFAULT_MAX_TEXT_BYTES,
        }
    }
}

/// Errors returned while constructing a bounded semantic document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A semantic profile exceeded a hard maximum.
    InvalidLimits {
        /// Rejected checked-limit request.
        error: LimitsError,
    },
    /// The supplied sheet sequence exceeds the selected bound.
    TooManySheets {
        /// Number of supplied sheets.
        actual: usize,
        /// Maximum accepted sheets.
        limit: usize,
    },
    /// A sheet does not carry its canonical position in the ordered sequence.
    InvalidSheetIndex {
        /// Position occupied by the sheet in the supplied sequence.
        expected: usize,
        /// Index stored by the sheet.
        actual: usize,
    },
    /// Two sheets use the same semantic name.
    DuplicateSheetName {
        /// Earlier sheet position using the name.
        first: usize,
        /// Later sheet position using the name.
        duplicate: usize,
    },
    /// The table aggregate exceeds the selected bound.
    TooManyTables {
        /// Number of supplied tables.
        actual: usize,
        /// Maximum accepted tables.
        limit: usize,
    },
    /// The materialized-cell aggregate exceeds the selected bound.
    TooManyMaterializedCells {
        /// Number of supplied materialized cells.
        actual: usize,
        /// Maximum accepted materialized cells.
        limit: usize,
    },
    /// Semantic names, headers, and textual cell values exceed the budget.
    TextTooLarge {
        /// Observed or requested UTF-8 bytes.
        observed: usize,
        /// Maximum accepted UTF-8 bytes.
        limit: usize,
    },
    /// A bounded semantic allocation failed before publication.
    Allocation {
        /// Requested elements or bytes.
        amount: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidLimits { error } => error.fmt(formatter),
            Self::TooManySheets { actual, limit } => write!(
                formatter,
                "Numbers document contains {actual} sheets; maximum is {limit}"
            ),
            Self::InvalidSheetIndex { expected, actual } => write!(
                formatter,
                "Numbers sheet index {actual} is not the expected index {expected}"
            ),
            Self::DuplicateSheetName { first, duplicate } => write!(
                formatter,
                "Numbers sheets {first} and {duplicate} have the same semantic name"
            ),
            Self::TooManyTables { actual, limit } => write!(
                formatter,
                "Numbers document contains {actual} tables; maximum is {limit}"
            ),
            Self::TooManyMaterializedCells { actual, limit } => write!(
                formatter,
                "Numbers document contains {actual} materialized cells; maximum is {limit}"
            ),
            Self::TextTooLarge { observed, limit } => write!(
                formatter,
                "Numbers document semantic text is {observed} bytes; maximum is {limit}"
            ),
            Self::Allocation { amount } => {
                write!(
                    formatter,
                    "Numbers document allocation failed for {amount} units"
                )
            },
        }
    }
}

impl std::error::Error for Error {}

/// Result type for bounded Numbers semantic construction.
pub type Result<T> = std::result::Result<T, Error>;

/// An immutable, archive-free Numbers document snapshot.
///
/// The document owns only semantic [`Sheet`] values. Its hidden state is
/// reference counted so cloning or taking a snapshot never copies the sheet
/// or table storage. Native archives, protobuf values, package entries, and
/// physical object identifiers are intentionally outside this API.
#[derive(Debug, Clone)]
pub struct Document {
    state: Arc<State>,
}

#[derive(Debug)]
struct State {
    sheets: Arc<[Sheet]>,
    plain_text_len: usize,
    diagnostics: Option<SourceDiagnostics>,
}

impl Document {
    /// Open a complete Numbers package or an app-authored package directory.
    ///
    /// The source is captured once and eagerly projected into this immutable,
    /// archive-free snapshot. Exact source bytes, media, previews, unsupported
    /// source items, and implementation identifiers are not retained. Use
    /// [`crate::Package`] when exact complete-package preservation or editing
    /// is required.
    ///
    /// # Errors
    ///
    /// Returns a typed error if source capture, format classification, or
    /// semantic projection fails or exceeds a checked resource ceiling.
    pub fn open(path: impl AsRef<Path>) -> std::result::Result<Self, ReadError> {
        Self::open_with_options(path, DocumentReadOptions::default())
    }

    /// Open an archive-free Numbers snapshot under explicit resource profiles.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::open`].
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: DocumentReadOptions,
    ) -> std::result::Result<Self, ReadError> {
        let source_limits = options
            .source()
            .detector_limits()
            .map_err(map_detection_error)?;
        let source = litchi_iwa_detect::PreparedSource::__from_path_with_numbers_semantics(
            path,
            source_limits,
        )
        .map_err(map_detection_error)?
        .ok_or(ReadError::InvalidSource)?;
        Self::from_prepared_source(source, options.semantic())
    }

    /// Decode borrowed packaged Numbers bytes into an archive-free snapshot.
    ///
    /// # Errors
    ///
    /// Returns a typed capture, format, or semantic projection failure.
    pub fn from_bytes(bytes: &[u8]) -> std::result::Result<Self, ReadError> {
        Self::from_bytes_with_options(bytes, DocumentReadOptions::default())
    }

    /// Decode borrowed package bytes under explicit resource profiles.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::from_bytes`].
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: DocumentReadOptions,
    ) -> std::result::Result<Self, ReadError> {
        let source_limits = options
            .source()
            .detector_limits()
            .map_err(map_detection_error)?;
        let source = litchi_iwa_detect::PreparedSource::__from_bytes_with_numbers_semantics(
            bytes,
            source_limits,
        )
        .map_err(map_detection_error)?
        .ok_or(ReadError::InvalidSource)?;
        Self::from_prepared_source(source, options.semantic())
    }

    /// Decode already-shared immutable package bytes without copying them at
    /// the capture boundary.
    ///
    /// The source allocation is released after eager semantic projection. Use
    /// [`crate::Package`] when exact bytes must remain authoritative.
    ///
    /// # Errors
    ///
    /// Returns a typed capture, format, or semantic projection failure.
    pub fn from_shared_bytes(bytes: Arc<[u8]>) -> std::result::Result<Self, ReadError> {
        Self::from_shared_bytes_with_options(bytes, DocumentReadOptions::default())
    }

    /// Decode shared immutable package bytes under explicit resource profiles.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::from_shared_bytes`].
    pub fn from_shared_bytes_with_options(
        bytes: Arc<[u8]>,
        options: DocumentReadOptions,
    ) -> std::result::Result<Self, ReadError> {
        let source_limits = options
            .source()
            .detector_limits()
            .map_err(map_detection_error)?;
        let source = litchi_iwa_detect::PreparedSource::__from_shared_bytes_with_numbers_semantics(
            bytes,
            source_limits,
        )
        .map_err(map_detection_error)?
        .ok_or(ReadError::InvalidSource)?;
        Self::from_prepared_source(source, options.semantic())
    }

    fn from_prepared_source(
        source: litchi_iwa_detect::PreparedSource,
        semantic: Limits,
    ) -> std::result::Result<Self, ReadError> {
        if source.format() != litchi_iwa_detect::Format::Numbers {
            return Err(ReadError::NotNumbers);
        }
        crate::package::semantic_document_from_prepared_source(source, semantic)
            .map_err(map_package_read_error)
    }

    /// Build a document from sheets in source order.
    ///
    /// # Errors
    ///
    /// Returns [`Error::TooManySheets`] when the hard semantic bound is
    /// exceeded, or [`Error::InvalidSheetIndex`] when a sheet is not numbered
    /// by its zero-based position in the supplied sequence.
    pub fn from_sheets(sheets: Vec<Sheet>) -> Result<Self> {
        Self::from_sheets_with_limits(sheets, Limits::default())
    }

    /// Build a document under a caller-selected sheet-count budget.
    ///
    /// The package-independent hard cap [`MAX_SHEETS`] cannot be relaxed by a
    /// caller. The input vector is consumed without rebuilding its sheet
    /// values when construction succeeds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::TooManySheets`] when the supplied count exceeds either
    /// the caller budget or the hard semantic cap, or
    /// [`Error::InvalidSheetIndex`] when a sheet is not numbered by its
    /// zero-based position in the supplied sequence.
    pub fn from_sheets_with_max_sheets(sheets: Vec<Sheet>, max_sheets: usize) -> Result<Self> {
        let limits = Limits::new(
            max_sheets,
            MAX_TABLES,
            MAX_MATERIALIZED_CELLS,
            DEFAULT_MAX_TEXT_BYTES,
        )
        .map_err(|error| Error::InvalidLimits { error })?;
        Self::from_sheets_with_limits(sheets, limits)
    }

    /// Build a document under explicit finite semantic budgets.
    ///
    /// The source vector is consumed without cloning its sheet or table
    /// values. Validation runs before the vector is moved into one shared
    /// immutable allocation, so a rejected archive cannot publish a partial
    /// semantic snapshot.
    ///
    /// # Errors
    ///
    /// Returns a typed error when any count, index, name, or textual-data
    /// budget is exceeded.
    pub fn from_sheets_with_limits(sheets: Vec<Sheet>, limits: Limits) -> Result<Self> {
        let max_sheets = limits.max_sheets.min(MAX_SHEETS);
        let max_tables = limits.max_tables.min(MAX_TABLES);
        let max_materialized_cells = limits.max_materialized_cells.min(MAX_MATERIALIZED_CELLS);
        let max_text_bytes = limits.max_text_bytes.min(DEFAULT_MAX_TEXT_BYTES);

        if sheets.len() > max_sheets {
            return Err(Error::TooManySheets {
                actual: sheets.len(),
                limit: max_sheets,
            });
        }

        let mut names = HashSet::new();
        names
            .try_reserve(sheets.len())
            .map_err(|_allocation| Error::Allocation {
                amount: sheets.len(),
            })?;
        let mut table_count = 0usize;
        let mut materialized_cell_count = 0usize;
        for (expected, sheet) in sheets.iter().enumerate() {
            if sheet.index() != expected {
                return Err(Error::InvalidSheetIndex {
                    expected,
                    actual: sheet.index(),
                });
            }
            if !names.insert(sheet.name()) {
                let first = sheets[..expected]
                    .iter()
                    .position(|previous| previous.name() == sheet.name())
                    .unwrap_or(expected);
                return Err(Error::DuplicateSheetName {
                    first,
                    duplicate: expected,
                });
            }

            for table in sheet.tables() {
                table_count = table_count.checked_add(1).ok_or(Error::TooManyTables {
                    actual: usize::MAX,
                    limit: max_tables,
                })?;
                if table_count > max_tables {
                    return Err(Error::TooManyTables {
                        actual: table_count,
                        limit: max_tables,
                    });
                }

                materialized_cell_count = materialized_cell_count
                    .checked_add(table.cell_count())
                    .ok_or(Error::TooManyMaterializedCells {
                    actual: usize::MAX,
                    limit: max_materialized_cells,
                })?;
                if materialized_cell_count > max_materialized_cells {
                    return Err(Error::TooManyMaterializedCells {
                        actual: materialized_cell_count,
                        limit: max_materialized_cells,
                    });
                }
            }
        }

        let plain_text_len = checked_plain_text_len(&sheets, max_text_bytes)?;
        Ok(Self {
            state: Arc::new(State {
                sheets: Arc::from(sheets.into_boxed_slice()),
                plain_text_len,
                diagnostics: None,
            }),
        })
    }

    pub(crate) fn from_source(
        mut document: Self,
        metadata: litchi_core::Metadata,
        stats: Stats,
    ) -> Self {
        if let Some(state) = Arc::get_mut(&mut document.state) {
            state.diagnostics = Some(SourceDiagnostics { metadata, stats });
            return document;
        }
        Self {
            state: Arc::new(State {
                sheets: Arc::clone(&document.state.sheets),
                plain_text_len: document.state.plain_text_len,
                diagnostics: Some(SourceDiagnostics { metadata, stats }),
            }),
        }
    }

    /// Capture another cheap handle to the same immutable snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Borrow all sheets in stable source order.
    #[must_use]
    pub fn sheets(&self) -> &[Sheet] {
        &self.state.sheets
    }

    /// Clone the shared sheet allocation without cloning any semantic values.
    #[must_use]
    pub fn shared_sheets(&self) -> Arc<[Sheet]> {
        Arc::clone(&self.state.sheets)
    }

    /// Select a sheet by its exact visible name or checked zero-based position.
    ///
    /// Names and positions identify the immutable semantic snapshot; native
    /// object identifiers are not part of this lookup boundary. Missing names
    /// and out-of-range positions return `Ok(None)`.
    ///
    /// ```rust,ignore
    /// let summary = document.sheet("Summary")?;
    /// let first = document.sheet(0)?;
    /// ```
    ///
    /// # Errors
    ///
    /// Validated immutable documents currently make this lookup infallible.
    /// The result keeps the selector API compatible with future validation
    /// rules that may need to report an ambiguous or invalid selector.
    pub fn sheet<'a, S>(&self, selector: S) -> Result<Option<&Sheet>>
    where
        S: Into<SheetSelector<'a>>,
    {
        match selector.into() {
            SheetSelector::Name(name) => {
                Ok(self.state.sheets.iter().find(|sheet| sheet.name() == name))
            },
            SheetSelector::Index(index) => Ok(self.state.sheets.get(index)),
        }
    }

    /// Return the number of semantic sheets.
    #[must_use]
    pub fn sheet_count(&self) -> usize {
        self.state.sheets.len()
    }

    /// Return whether the document contains no semantic sheets.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.state.sheets.is_empty()
    }

    /// Render rooted workbook values in deterministic semantic order.
    ///
    /// Sheets and tables retain source order; materialized cells retain
    /// row-major order. Non-empty sheet names, table names, and cell display
    /// values are emitted with one newline between values. Headers, missing
    /// cells, auxiliary drawings, comments, charts, templates, and source
    /// storages are outside this rooted workbook projection.
    ///
    /// # Errors
    ///
    /// Returns a content-free allocation failure if the exact validated output
    /// capacity cannot be reserved.
    pub fn plain_text(&self) -> std::result::Result<String, ReadError> {
        let mut text = String::new();
        text.try_reserve_exact(self.state.plain_text_len)
            .map_err(|_allocation| ReadError::Allocation {
                amount: self.state.plain_text_len,
            })?;
        append_plain_text(&self.state.sheets, &mut text);
        debug_assert_eq!(text.len(), self.state.plain_text_len);
        Ok(text)
    }

    /// Return the exact UTF-8 length of [`Self::plain_text`].
    #[must_use]
    pub fn text_len(&self) -> usize {
        self.state.plain_text_len
    }

    /// Return source diagnostics retained during bounded ingress.
    ///
    /// Documents built from semantic values, including documents borrowed from
    /// [`crate::Package`], have no physical source diagnostics and return
    /// `None`.
    #[must_use]
    pub fn stats(&self) -> Option<Stats> {
        self.state.diagnostics.as_ref().map(|value| value.stats)
    }

    /// Borrow metadata captured from a validated Numbers source.
    ///
    /// Only the three canonical document authorities are interpreted. Values
    /// built from semantic sheets, including [`crate::Package::document`],
    /// have no source diagnostics and return `None`.
    #[must_use]
    pub fn metadata(&self) -> Option<&litchi_core::Metadata> {
        self.state.diagnostics.as_ref().map(|value| &value.metadata)
    }

    /// Revalidate the detached semantic snapshot without source bytes.
    ///
    /// # Errors
    ///
    /// Returns a typed semantic error if retained values violate document
    /// invariants.
    pub fn validate(&self) -> Result<()> {
        validate_sheets(&self.state.sheets, Limits::default())
    }
}

fn checked_text_add(current: usize, added: usize, limit: usize) -> Result<usize> {
    let total = current.checked_add(added).ok_or(Error::TextTooLarge {
        observed: usize::MAX,
        limit,
    })?;
    if total > limit {
        return Err(Error::TextTooLarge {
            observed: total,
            limit,
        });
    }
    Ok(total)
}

fn checked_plain_text_len(sheets: &[Sheet], limit: usize) -> Result<usize> {
    let mut length = 0usize;
    let mut lines = 0usize;
    for sheet in sheets {
        if !sheet.name().is_empty() {
            charge_plain_text_line(&mut length, &mut lines, sheet.name().len(), limit)?;
        }
        for table in sheet.tables() {
            if !table.name().is_empty() {
                charge_plain_text_line(&mut length, &mut lines, table.name().len(), limit)?;
            }
            for cell in table.iter_cells() {
                if cell.value().is_empty() {
                    continue;
                }
                let value_len = value_text_len(cell.value());
                if value_len == 0 {
                    continue;
                }
                charge_plain_text_line(&mut length, &mut lines, value_len, limit)?;
            }
        }
    }
    Ok(length)
}

fn charge_plain_text_line(
    length: &mut usize,
    lines: &mut usize,
    value_len: usize,
    limit: usize,
) -> Result<()> {
    *length = checked_text_add(*length, usize::from(*lines != 0), limit)?;
    *length = checked_text_add(*length, value_len, limit)?;
    *lines = lines.saturating_add(1);
    Ok(())
}

fn value_text_len(value: &Value) -> usize {
    match value {
        Value::Empty => 0,
        Value::Text(value) | Value::Formula(value) => value.len(),
        Value::Number(value) | Value::Date(value) | Value::Duration(value) => {
            let mut counter = CountingWriter::default();
            let _ = write!(&mut counter, "{}", value.get());
            counter.bytes
        },
        Value::Boolean(true) => 4,
        Value::Boolean(false) => 5,
        Value::Error(value) => "ERROR: ".len().saturating_add(value.len()),
    }
}

#[derive(Default)]
struct CountingWriter {
    bytes: usize,
}

impl fmt::Write for CountingWriter {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.bytes = self.bytes.checked_add(value.len()).ok_or(fmt::Error)?;
        Ok(())
    }
}

fn append_plain_text(sheets: &[Sheet], output: &mut String) {
    let mut has_line = false;
    for sheet in sheets {
        if !sheet.name().is_empty() {
            append_plain_text_line(output, &mut has_line, sheet.name());
        }
        for table in sheet.tables() {
            if !table.name().is_empty() {
                append_plain_text_line(output, &mut has_line, table.name());
            }
            for cell in table.iter_cells() {
                if cell.value().is_empty() {
                    continue;
                }
                if value_text_len(cell.value()) == 0 {
                    continue;
                }
                if has_line {
                    output.push('\n');
                }
                match cell.value() {
                    Value::Empty => {},
                    Value::Text(value) | Value::Formula(value) => output.push_str(value),
                    Value::Number(value) | Value::Date(value) | Value::Duration(value) => {
                        let _ = write!(output, "{}", value.get());
                    },
                    Value::Boolean(value) => {
                        output.push_str(if *value { "true" } else { "false" });
                    },
                    Value::Error(value) => {
                        output.push_str("ERROR: ");
                        output.push_str(value);
                    },
                }
                has_line = true;
            }
        }
    }
}

fn append_plain_text_line(output: &mut String, has_line: &mut bool, value: &str) {
    if *has_line {
        output.push('\n');
    }
    output.push_str(value);
    *has_line = true;
}

fn validate_sheets(sheets: &[Sheet], limits: Limits) -> Result<()> {
    let max_sheets = limits.max_sheets.min(MAX_SHEETS);
    let max_tables = limits.max_tables.min(MAX_TABLES);
    let max_cells = limits.max_materialized_cells.min(MAX_MATERIALIZED_CELLS);
    let max_text = limits.max_text_bytes.min(DEFAULT_MAX_TEXT_BYTES);
    if sheets.len() > max_sheets {
        return Err(Error::TooManySheets {
            actual: sheets.len(),
            limit: max_sheets,
        });
    }
    let mut table_count = 0usize;
    let mut cell_count = 0usize;
    for (expected, sheet) in sheets.iter().enumerate() {
        if sheet.index() != expected {
            return Err(Error::InvalidSheetIndex {
                expected,
                actual: sheet.index(),
            });
        }
        if let Some(first) = sheets[..expected]
            .iter()
            .position(|previous| previous.name() == sheet.name())
        {
            return Err(Error::DuplicateSheetName {
                first,
                duplicate: expected,
            });
        }
        for table in sheet.tables() {
            table_count = table_count.checked_add(1).ok_or(Error::TooManyTables {
                actual: usize::MAX,
                limit: max_tables,
            })?;
            if table_count > max_tables {
                return Err(Error::TooManyTables {
                    actual: table_count,
                    limit: max_tables,
                });
            }
            cell_count = cell_count.checked_add(table.cell_count()).ok_or(
                Error::TooManyMaterializedCells {
                    actual: usize::MAX,
                    limit: max_cells,
                },
            )?;
            if cell_count > max_cells {
                return Err(Error::TooManyMaterializedCells {
                    actual: cell_count,
                    limit: max_cells,
                });
            }
        }
    }
    let _ = checked_plain_text_len(sheets, max_text)?;
    Ok(())
}

fn check_source_limit(
    kind: DocumentSourceLimitKind,
    value: u64,
    maximum: u64,
) -> std::result::Result<(), DocumentSourceLimitsError> {
    if value == 0 || value > maximum {
        return Err(DocumentSourceLimitsError {
            kind,
            value,
            maximum,
        });
    }
    Ok(())
}

fn map_package_read_error(error: crate::PackageError) -> ReadError {
    match error {
        crate::PackageError::Io(error) => ReadError::Io {
            kind: io_kind(error.kind()),
        },
        crate::PackageError::Archive(error) => map_archive_error(error),
        crate::PackageError::Detection(error) => map_detection_error(error),
        crate::PackageError::NotNumbers => ReadError::NotNumbers,
        crate::PackageError::MalformedPayload { .. }
        | crate::PackageError::InvalidFormat(_)
        | crate::PackageError::ParseError(_) => ReadError::InvalidFormat,
        crate::PackageError::Common(error) => map_common_error(error),
        crate::PackageError::Semantic(error) => map_semantic_error(error),
        crate::PackageError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => ReadError::Limit {
            kind: semantic_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        crate::PackageError::InputTooLarge { observed, maximum } => ReadError::Limit {
            kind: ReadLimitKind::InputBytes,
            observed,
            maximum,
        },
    }
}

fn map_semantic_error(error: Error) -> ReadError {
    match error {
        Error::InvalidLimits { .. } => ReadError::InvalidSource,
        Error::TooManySheets { actual, limit } => ReadError::Limit {
            kind: ReadLimitKind::Sheets,
            observed: usize_u64(actual),
            maximum: usize_u64(limit),
        },
        Error::TooManyTables { actual, limit } => ReadError::Limit {
            kind: ReadLimitKind::Tables,
            observed: usize_u64(actual),
            maximum: usize_u64(limit),
        },
        Error::TooManyMaterializedCells { actual, limit } => ReadError::Limit {
            kind: ReadLimitKind::Cells,
            observed: usize_u64(actual),
            maximum: usize_u64(limit),
        },
        Error::TextTooLarge { observed, limit } => ReadError::Limit {
            kind: ReadLimitKind::TextBytes,
            observed: usize_u64(observed),
            maximum: usize_u64(limit),
        },
        Error::Allocation { amount } => ReadError::Allocation { amount },
        Error::InvalidSheetIndex { .. } | Error::DuplicateSheetName { .. } => {
            ReadError::InvalidFormat
        },
    }
}

fn map_detection_error(error: litchi_iwa_detect::Error) -> ReadError {
    match error {
        litchi_iwa_detect::Error::Io(error) => ReadError::Io {
            kind: io_kind(error.kind()),
        },
        litchi_iwa_detect::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => ReadError::Limit {
            kind: detection_limit_kind(kind),
            observed,
            maximum,
        },
        litchi_iwa_detect::Error::Allocation { amount } => ReadError::Allocation { amount },
        litchi_iwa_detect::Error::SourceChanged => ReadError::InvalidSource,
        litchi_iwa_detect::Error::IwaCore(error) => map_core_error(error),
        litchi_iwa_detect::Error::IwaCommon(error) => map_common_error(error),
        litchi_iwa_detect::Error::InvalidFormat(_)
        | litchi_iwa_detect::Error::Archive(_)
        | litchi_iwa_detect::Error::InvalidLimits
        | litchi_iwa_detect::Error::Encrypted => ReadError::InvalidFormat,
        _ => ReadError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> ReadError {
    match error {
        litchi_iwa_archive::Error::Io(error) => ReadError::Io {
            kind: io_kind(error.kind()),
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => ReadError::Limit {
            kind: archive_limit_kind(kind),
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => ReadError::Allocation { amount },
        litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. } => ReadError::InvalidSource,
        litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => ReadError::InvalidFormat,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> ReadError {
    match error {
        litchi_iwa_core::Error::Io(error) => ReadError::Io {
            kind: io_kind(error.kind()),
        },
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => ReadError::Limit {
            kind: core_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            ReadError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Snappy { .. } => ReadError::InvalidFormat,
    }
}

fn map_common_error(error: litchi_iwa_common::Error) -> ReadError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => ReadError::Limit {
            kind: common_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => ReadError::Allocation { amount },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => ReadError::InvalidFormat,
    }
}

fn semantic_limit_kind(kind: crate::SemanticLimitKind) -> ReadLimitKind {
    match kind {
        crate::SemanticLimitKind::Sheets => ReadLimitKind::Sheets,
        crate::SemanticLimitKind::Tables => ReadLimitKind::Tables,
        crate::SemanticLimitKind::MaterializedCells => ReadLimitKind::Cells,
        crate::SemanticLimitKind::OutputTextBytes | crate::SemanticLimitKind::TextBytes => {
            ReadLimitKind::TextBytes
        },
        crate::SemanticLimitKind::FormulaRenderDepth | crate::SemanticLimitKind::FormulaDepth => {
            ReadLimitKind::PayloadNesting
        },
        crate::SemanticLimitKind::Objects | crate::SemanticLimitKind::References => {
            ReadLimitKind::PayloadFields
        },
        crate::SemanticLimitKind::FormulaWireBytes => ReadLimitKind::PayloadBytes,
        crate::SemanticLimitKind::FormulaRenderWork | crate::SemanticLimitKind::FormulaWork => {
            ReadLimitKind::PayloadWork
        },
    }
}

fn detection_limit_kind(kind: litchi_iwa_detect::LimitKind) -> ReadLimitKind {
    match kind {
        litchi_iwa_detect::LimitKind::InputBytes => ReadLimitKind::InputBytes,
        litchi_iwa_detect::LimitKind::Entries => ReadLimitKind::Entries,
        litchi_iwa_detect::LimitKind::OutputBytes
        | litchi_iwa_detect::LimitKind::MemberNameBytes
        | litchi_iwa_detect::LimitKind::MetadataBytes
        | litchi_iwa_detect::LimitKind::CompressedEntryBytes
        | litchi_iwa_detect::LimitKind::EntryBytes
        | litchi_iwa_detect::LimitKind::TotalBytes
        | litchi_iwa_detect::LimitKind::IwaStreamBytes
        | litchi_iwa_detect::LimitKind::IwaTotalBytes => ReadLimitKind::PayloadBytes,
        litchi_iwa_detect::LimitKind::IwaObjects => ReadLimitKind::PayloadFields,
        litchi_iwa_detect::LimitKind::IwaFields => ReadLimitKind::PayloadFields,
        litchi_iwa_detect::LimitKind::IwaNesting => ReadLimitKind::PayloadNesting,
        litchi_iwa_detect::LimitKind::IwaWork => ReadLimitKind::PayloadWork,
        _ => ReadLimitKind::Other,
    }
}

fn archive_limit_kind(kind: litchi_iwa_archive::LimitKind) -> ReadLimitKind {
    match kind {
        litchi_iwa_archive::LimitKind::InputBytes => ReadLimitKind::InputBytes,
        litchi_iwa_archive::LimitKind::Entries => ReadLimitKind::Entries,
        litchi_iwa_archive::LimitKind::OutputBytes
        | litchi_iwa_archive::LimitKind::MemberNameBytes
        | litchi_iwa_archive::LimitKind::MetadataBytes
        | litchi_iwa_archive::LimitKind::CompressedEntryBytes
        | litchi_iwa_archive::LimitKind::EntryBytes
        | litchi_iwa_archive::LimitKind::TotalBytes
        | litchi_iwa_archive::LimitKind::IwaStreamBytes
        | litchi_iwa_archive::LimitKind::IwaTotalBytes => ReadLimitKind::PayloadBytes,
    }
}

fn core_limit_kind(kind: litchi_iwa_core::LimitKind) -> ReadLimitKind {
    match kind {
        litchi_iwa_core::LimitKind::Objects
        | litchi_iwa_core::LimitKind::Messages
        | litchi_iwa_core::LimitKind::MessagesPerObject
        | litchi_iwa_core::LimitKind::HeaderFields
        | litchi_iwa_core::LimitKind::MetadataItems
        | litchi_iwa_core::LimitKind::SnappyFrames => ReadLimitKind::PayloadFields,
        litchi_iwa_core::LimitKind::HeaderNesting => ReadLimitKind::PayloadNesting,
        _ => ReadLimitKind::PayloadBytes,
    }
}

fn common_limit_kind(kind: litchi_iwa_common::LimitKind) -> ReadLimitKind {
    match kind {
        litchi_iwa_common::LimitKind::InputBytes | litchi_iwa_common::LimitKind::OutputBytes => {
            ReadLimitKind::PayloadBytes
        },
        litchi_iwa_common::LimitKind::Fields => ReadLimitKind::PayloadFields,
        litchi_iwa_common::LimitKind::Nesting => ReadLimitKind::PayloadNesting,
        litchi_iwa_common::LimitKind::MaterializedCells => ReadLimitKind::Cells,
        litchi_iwa_common::LimitKind::RewriteWork
        | litchi_iwa_common::LimitKind::TableRows
        | litchi_iwa_common::LimitKind::TableColumns
        | litchi_iwa_common::LimitKind::TableCells => ReadLimitKind::PayloadWork,
    }
}

fn io_kind(kind: std::io::ErrorKind) -> IoKind {
    match kind {
        std::io::ErrorKind::NotFound => IoKind::NotFound,
        std::io::ErrorKind::PermissionDenied => IoKind::PermissionDenied,
        std::io::ErrorKind::AlreadyExists => IoKind::AlreadyExists,
        std::io::ErrorKind::InvalidInput => IoKind::InvalidInput,
        std::io::ErrorKind::InvalidData => IoKind::InvalidData,
        std::io::ErrorKind::TimedOut => IoKind::TimedOut,
        std::io::ErrorKind::Interrupted => IoKind::Interrupted,
        std::io::ErrorKind::UnexpectedEof => IoKind::UnexpectedEof,
        _ => IoKind::Other,
    }
}

fn usize_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn empty_document_is_a_valid_bounded_snapshot() {
        let document = Document::from_sheets(Vec::new())
            .unwrap_or_else(|error| panic!("empty document should be valid: {error}"));

        assert_send_sync::<Document>();
        assert!(document.is_empty());
        assert_eq!(document.sheet_count(), 0);
        assert!(document.sheets().is_empty());
        assert_eq!(document.plain_text().unwrap(), "");
        assert_eq!(document.text_len(), 0);
        assert_eq!(document.stats(), None);
        assert!(document.metadata().is_none());
        assert!(
            document
                .sheet(0)
                .unwrap_or_else(|error| panic!("empty document lookup failed: {error}"))
                .is_none()
        );
    }

    #[test]
    fn construction_checks_budget_and_canonical_order() {
        let too_many = Document::from_sheets_with_max_sheets(vec![Sheet::new("Sheet 1", 0)], 0);
        assert!(matches!(
            too_many,
            Err(Error::TooManySheets {
                actual: 1,
                limit: 0,
            })
        ));

        let invalid = Document::from_sheets(vec![Sheet::new("Sheet 2", 1)]);
        assert!(matches!(
            invalid,
            Err(Error::InvalidSheetIndex {
                expected: 0,
                actual: 1,
            })
        ));
    }

    #[test]
    fn clones_share_ordered_semantic_storage() {
        let document =
            Document::from_sheets(vec![Sheet::new("Sheet 1", 0), Sheet::new("Sheet 2", 1)])
                .unwrap_or_else(|error| panic!("document should be valid: {error}"));
        let snapshot = document.snapshot();

        assert!(Arc::ptr_eq(&document.state.sheets, &snapshot.state.sheets));
        assert_eq!(snapshot.sheet_count(), 2);
        assert_eq!(
            snapshot
                .sheet(0)
                .unwrap_or_else(|error| panic!("index lookup failed: {error}"))
                .map(Sheet::name),
            Some("Sheet 1")
        );
        assert_eq!(
            snapshot
                .sheet(1)
                .unwrap_or_else(|error| panic!("index lookup failed: {error}"))
                .map(Sheet::name),
            Some("Sheet 2")
        );
        assert!(
            snapshot
                .sheet(2)
                .unwrap_or_else(|error| panic!("index lookup failed: {error}"))
                .is_none()
        );
        assert_eq!(
            snapshot
                .sheet("Sheet 2")
                .unwrap_or_else(|error| panic!("name lookup failed: {error}"))
                .map(Sheet::index),
            Some(1)
        );
    }

    #[test]
    fn sheet_lookup_is_selector_first_and_returns_none_when_missing() {
        let document =
            Document::from_sheets(vec![Sheet::new("Summary", 0), Sheet::new("Archive", 1)])
                .unwrap_or_else(|error| panic!("document should be valid: {error}"));

        assert_eq!(
            document
                .sheet("Summary")
                .unwrap_or_else(|error| panic!("name lookup failed: {error}"))
                .map(Sheet::index),
            Some(0)
        );
        assert_eq!(
            document
                .sheet(1)
                .unwrap_or_else(|error| panic!("index lookup failed: {error}"))
                .map(Sheet::name),
            Some("Archive")
        );
        assert!(
            document
                .sheet("Missing")
                .unwrap_or_else(|error| panic!("name lookup failed: {error}"))
                .is_none()
        );
        assert!(
            document
                .sheet(2)
                .unwrap_or_else(|error| panic!("index lookup failed: {error}"))
                .is_none()
        );
    }

    #[test]
    fn construction_rejects_duplicate_names_and_aggregate_budgets() {
        let duplicate =
            Document::from_sheets(vec![Sheet::new("Summary", 0), Sheet::new("Summary", 1)]);
        assert!(matches!(
            duplicate,
            Err(Error::DuplicateSheetName {
                first: 0,
                duplicate: 1,
            })
        ));

        let mut table = crate::table::Builder::new("Data", crate::Dimensions::new(1, 1));
        assert!(
            table
                .set(crate::Position::new(0, 0), Value::Text("value".to_owned()))
                .is_ok()
        );
        let table = table
            .finish()
            .unwrap_or_else(|error| panic!("table should be valid: {error}"));
        let mut sheet = crate::sheet::Builder::new("Summary", 0);
        assert!(sheet.push_table(table).is_ok());
        let sheet = sheet.finish();

        let table_limit = Limits::new(1, 0, MAX_MATERIALIZED_CELLS, DEFAULT_MAX_TEXT_BYTES)
            .unwrap_or_else(|error| panic!("valid zero table limit rejected: {error}"));
        assert!(matches!(
            Document::from_sheets_with_limits(vec![sheet.clone()], table_limit),
            Err(Error::TooManyTables {
                actual: 1,
                limit: 0,
            })
        ));

        let text_limit = Limits::new(1, MAX_TABLES, MAX_MATERIALIZED_CELLS, 4)
            .unwrap_or_else(|error| panic!("valid text limit rejected: {error}"));
        assert!(matches!(
            Document::from_sheets_with_limits(vec![sheet], text_limit),
            Err(Error::TextTooLarge {
                observed: _,
                limit: 4,
            })
        ));
    }
}
