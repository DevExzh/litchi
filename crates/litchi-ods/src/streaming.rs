//! Bounded sequential creation of one-sheet ODS scalar documents.
//!
//! This module exposes a narrow fresh-document authoring surface. It consumes
//! rows once, validates scalar values and XML 1.0 text, emits a fixed `Sheet1`
//! envelope through the common generated-XML seam, and writes to a
//! caller-owned non-seeking sink. The public grammar accepts arbitrary bounded
//! scalar rows; benchmark fixtures may impose their own four-cell shape.
//!
//! The ODS provider owns row grammar, execution budgeting, and sink progress.
//! The common ODF writer owns lexical XML auditing, generated-member admission,
//! ZIP framing, manifest bookkeeping, and physical publication.

use std::{borrow::Cow, error::Error as StdError, fmt, io::Write};

use litchi_core::{Error, ExecutionContext, ExecutionError, Resource};
use litchi_odf_common::core::{
    GeneratedXmlEnvelope, GeneratedXmlLimits, PackageWriter, PackageWriterError,
    PackageWriterLimits,
};

const ODS_MIME: &str = "application/vnd.oasis.opendocument.spreadsheet";
const CONTENT_PATH: &str = "content.xml";
const CONTENT_MEDIA_TYPE: &str = "text/xml";

// Sheet1 is part of the frozen one-sheet creation contract. Keeping the
// envelope static avoids a provider-side shell allocation; the common
// envelope constructor owns its checked copy for publication.
const CONTENT_PREFIX: &[u8] = b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.3\"><office:body><office:spreadsheet><table:table table:name=\"Sheet1\">";
const CONTENT_SUFFIX: &[u8] =
    b"</table:table></office:spreadsheet></office:body></office:document-content>";

// Reuse the already-admitted worksheet safety constants so this narrow
// streaming surface cannot create a document outside the ordinary ODS
// validation envelope.
const MAX_PHYSICAL_ROWS: usize = crate::worksheet::validation::MAX_PHYSICAL_RUNS;
const DEFAULT_MAX_CELLS: usize = crate::worksheet::validation::MAX_LOGICAL_CELLS;
const DEFAULT_MAX_CELLS_PER_ROW: usize = crate::worksheet::validation::MAX_PHYSICAL_RUNS;
// The streaming surface narrows the ordinary 16 MiB worksheet field ceiling
// to the accepted 1 MiB per-cell text budget.
const MAX_TEXT_BYTES: usize = 1 << 20;
const MAX_CONTENT_XML_BYTES: usize = crate::worksheet::validation::MAX_CONTENT_XML_BYTES;
const DEFAULT_MAX_ROWS: usize = MAX_PHYSICAL_ROWS;
const DEFAULT_MAX_TEXT_BYTES: usize = MAX_TEXT_BYTES;
const DEFAULT_MAX_ROW_XML_BYTES: usize = 4096;
const DEFAULT_MAX_CONTENT_XML_BYTES: usize = 32 << 20;
const DEFAULT_MAX_OUTPUT_BYTES: u64 = 512 << 20;
const COMMON_METADATA_RESERVATION: usize = 64 * 1024;

/// A scalar cell accepted by the streaming worksheet grammar.
///
/// `Text` is a `Cow` so a caller can pass a borrowed field from a row source
/// without copying it. The common publication buffer is the only retained XML
/// window; its contents are cleared and reused after each row while its fixed
/// capacity remains reserved until publication finishes.
#[derive(Debug, Clone, PartialEq)]
pub enum StreamingCell<'a> {
    /// A finite ODF floating-point value.
    Number(f64),
    /// A UTF-8 string whose XML 1.0 characters are checked before emission.
    Text(Cow<'a, str>),
    /// An ODF boolean value.
    Boolean(bool),
    /// An empty table cell.
    Empty,
}

/// Scalar XML audit limits exposed without leaking the XML-minifier type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XmlAuditLimits {
    max_bytes: usize,
    max_depth: usize,
    max_events: usize,
    max_attributes: usize,
    max_token_bytes: usize,
    max_text_bytes: usize,
}

impl XmlAuditLimits {
    /// Creates an explicit finite XML audit profile.
    pub fn new(
        max_bytes: usize,
        max_depth: usize,
        max_events: usize,
        max_attributes: usize,
        max_token_bytes: usize,
        max_text_bytes: usize,
    ) -> Result<Self, StreamingError> {
        let limits = GeneratedXmlLimits::new(
            max_bytes,
            max_depth,
            max_events,
            max_attributes,
            max_token_bytes,
            max_text_bytes,
        )
        .map_err(|error| invalid(format!("invalid XML audit limits: {error}")))?;
        Ok(Self::from_common(limits))
    }

    /// Returns the aggregate byte ceiling.
    #[must_use]
    pub const fn max_bytes(self) -> usize {
        self.max_bytes
    }

    /// Returns the element-depth ceiling.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Returns the parser-event ceiling.
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.max_events
    }

    /// Returns the aggregate attribute ceiling.
    #[must_use]
    pub const fn max_attributes(self) -> usize {
        self.max_attributes
    }

    /// Returns the single-token ceiling.
    #[must_use]
    pub const fn max_token_bytes(self) -> usize {
        self.max_token_bytes
    }

    /// Returns the aggregate character-data ceiling.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    fn from_common(limits: GeneratedXmlLimits) -> Self {
        Self {
            max_bytes: limits.max_bytes(),
            max_depth: limits.max_depth(),
            max_events: limits.max_events(),
            max_attributes: limits.max_attributes(),
            max_token_bytes: limits.max_token_bytes(),
            max_text_bytes: limits.max_text_bytes(),
        }
    }

    fn into_common(self) -> Result<GeneratedXmlLimits, StreamingError> {
        // Construction already checked every field against the immutable
        // minifier ceilings.  Keep this conversion fallible in spirit by
        // preserving the checked constructor at every public boundary.
        GeneratedXmlLimits::new(
            self.max_bytes,
            self.max_depth,
            self.max_events,
            self.max_attributes,
            self.max_token_bytes,
            self.max_text_bytes,
        )
        .map_err(|error| invalid(format!("invalid XML audit limits: {error}")))
    }
}

impl Default for XmlAuditLimits {
    fn default() -> Self {
        Self::from_common(GeneratedXmlLimits::default())
    }
}

/// Provider-owned finite limits for the row stream and its XML fragments.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StreamingLimits {
    max_rows: usize,
    max_cells: usize,
    max_cells_per_row: usize,
    max_text_bytes: usize,
    max_row_xml_bytes: usize,
    max_content_xml_bytes: usize,
    max_output_bytes: u64,
    xml_audit: XmlAuditLimits,
}

impl StreamingLimits {
    /// Creates a checked provider limit profile.
    #[allow(clippy::too_many_arguments)]
    pub fn new(
        max_rows: usize,
        max_cells: usize,
        max_cells_per_row: usize,
        max_text_bytes: usize,
        max_row_xml_bytes: usize,
        max_content_xml_bytes: usize,
        max_output_bytes: u64,
        xml_audit: XmlAuditLimits,
    ) -> Result<Self, StreamingError> {
        if max_rows == 0 || max_rows > MAX_PHYSICAL_ROWS {
            return Err(invalid(format!(
                "max_rows must be in 1..={MAX_PHYSICAL_ROWS}"
            )));
        }
        if max_cells == 0 || max_cells > DEFAULT_MAX_CELLS {
            return Err(invalid(format!(
                "max_cells must be in 1..={DEFAULT_MAX_CELLS}"
            )));
        }
        if max_cells_per_row == 0 || max_cells_per_row > DEFAULT_MAX_CELLS_PER_ROW {
            return Err(invalid(format!(
                "max_cells_per_row must be in 1..={DEFAULT_MAX_CELLS_PER_ROW}"
            )));
        }
        if max_text_bytes == 0 || max_text_bytes > MAX_TEXT_BYTES {
            return Err(invalid(format!(
                "max_text_bytes must be in 1..={MAX_TEXT_BYTES}"
            )));
        }
        if max_row_xml_bytes == 0 {
            return Err(invalid("max_row_xml_bytes must be greater than zero"));
        }
        if max_content_xml_bytes == 0 || max_content_xml_bytes > MAX_CONTENT_XML_BYTES {
            return Err(invalid(format!(
                "max_content_xml_bytes must be in 1..={MAX_CONTENT_XML_BYTES}"
            )));
        }
        if max_row_xml_bytes > max_content_xml_bytes {
            return Err(invalid(
                "max_row_xml_bytes must not exceed max_content_xml_bytes",
            ));
        }
        if max_output_bytes < 22 {
            return Err(invalid(
                "max_output_bytes must fit the minimum empty ZIP end record",
            ));
        }
        if max_content_xml_bytes > xml_audit.max_bytes() {
            return Err(invalid(
                "max_content_xml_bytes must not exceed the XML byte audit limit",
            ));
        }
        if max_text_bytes > xml_audit.max_text_bytes() {
            return Err(invalid(
                "max_text_bytes must not exceed the XML aggregate text limit",
            ));
        }
        Ok(Self {
            max_rows,
            max_cells,
            max_cells_per_row,
            max_text_bytes,
            max_row_xml_bytes,
            max_content_xml_bytes,
            max_output_bytes,
            xml_audit,
        })
    }

    /// Narrows the XML audit profile while retaining the stream limits.
    pub fn with_xml_audit_limits(self, xml_audit: XmlAuditLimits) -> Result<Self, StreamingError> {
        Self::new(
            self.max_rows,
            self.max_cells,
            self.max_cells_per_row,
            self.max_text_bytes,
            self.max_row_xml_bytes,
            self.max_content_xml_bytes,
            self.max_output_bytes,
            xml_audit,
        )
    }

    /// Maximum rows accepted from the iterator.
    #[must_use]
    pub const fn max_rows(self) -> usize {
        self.max_rows
    }

    /// Maximum cells accepted across all rows.
    #[must_use]
    pub const fn max_cells(self) -> usize {
        self.max_cells
    }

    /// Maximum cells accepted in one row.
    #[must_use]
    pub const fn max_cells_per_row(self) -> usize {
        self.max_cells_per_row
    }

    /// Maximum UTF-8 bytes in one text value.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Maximum bytes in one emitted row fragment.
    #[must_use]
    pub const fn max_row_xml_bytes(self) -> usize {
        self.max_row_xml_bytes
    }

    /// Maximum authored bytes in `content.xml`.
    #[must_use]
    pub const fn max_content_xml_bytes(self) -> usize {
        self.max_content_xml_bytes
    }

    /// Maximum complete ZIP bytes accepted by the caller sink.
    #[must_use]
    pub const fn max_output_bytes(self) -> u64 {
        self.max_output_bytes
    }

    /// Retained `Resource::Memory` reservation for one streaming operation.
    ///
    /// This covers the reusable row-fragment window, two copies of the fixed
    /// XML envelope shell, and common metadata staging. It does not model all
    /// ZIP, XML-auditor, or allocator physical peak memory.
    #[must_use]
    pub fn required_memory_bytes(self) -> u64 {
        let shell_bytes = CONTENT_PREFIX
            .len()
            .checked_add(CONTENT_SUFFIX.len())
            .expect("fixed ODS envelope byte count must fit usize");
        let envelope_memory = shell_bytes
            .checked_mul(2)
            .expect("fixed ODS envelope memory count must fit usize");
        let retained_memory = self
            .max_row_xml_bytes
            .checked_add(envelope_memory)
            .and_then(|size| size.checked_add(COMMON_METADATA_RESERVATION))
            .expect("checked streaming limits must fit retained ODS memory");
        u64::try_from(retained_memory).expect("checked retained ODS memory must fit u64")
    }

    /// XML-minifier limits used by the common generated XML seam.
    #[must_use]
    pub const fn xml_audit(self) -> XmlAuditLimits {
        self.xml_audit
    }
}

impl Default for StreamingLimits {
    fn default() -> Self {
        // These constants are kept in the same order and below the immutable
        // XML/package ceilings.  The checked constructor remains the public
        // path for caller-selected profiles.
        Self {
            max_rows: DEFAULT_MAX_ROWS,
            max_cells: DEFAULT_MAX_CELLS,
            max_cells_per_row: DEFAULT_MAX_CELLS_PER_ROW,
            max_text_bytes: DEFAULT_MAX_TEXT_BYTES,
            max_row_xml_bytes: DEFAULT_MAX_ROW_XML_BYTES,
            max_content_xml_bytes: DEFAULT_MAX_CONTENT_XML_BYTES,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
            xml_audit: XmlAuditLimits::default(),
        }
    }
}

/// Counters returned after the provider has released all operation buffers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ScalarStreamReport {
    rows: usize,
    cells: usize,
    authored_content_xml_bytes: usize,
}

impl ScalarStreamReport {
    /// Number of accepted rows.
    #[must_use]
    pub const fn rows(self) -> usize {
        self.rows
    }

    /// Number of accepted cells.
    #[must_use]
    pub const fn cells(self) -> usize {
        self.cells
    }

    /// Authored `content.xml` bytes, before ZIP compression.
    #[must_use]
    pub const fn authored_content_xml_bytes(self) -> usize {
        self.authored_content_xml_bytes
    }
}

/// Format-neutral source category for an output failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PublicationFailureKind {
    /// Metadata, envelope, or other pre-publication refusal.
    Preflight,
    /// A row iterator or typed XML producer failed.
    Producer,
    /// The caller sink rejected bytes.
    Sink,
    /// ZIP framing or member finalization failed.
    Archive,
    /// A finite ZIP/XML/output limit was reached.
    Limit,
    /// Final archive manifest or end-record publication failed.
    Finalization,
}

/// A publication failure that preserves accepted sink progress without
/// exposing the concrete ZIP error enum through the ODS API.
#[derive(Debug)]
pub struct PublicationError {
    written: u64,
    kind: PublicationFailureKind,
    source: Box<dyn StdError + Send + Sync + 'static>,
}

impl PublicationError {
    /// Number of bytes accepted by the caller sink before failure.
    #[must_use]
    pub const fn written(&self) -> u64 {
        self.written
    }

    /// Format-neutral failure category.
    #[must_use]
    pub const fn kind(&self) -> PublicationFailureKind {
        self.kind
    }
}

impl fmt::Display for PublicationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "ODS publication failed after {} byte(s): {}",
            self.written, self.source
        )
    }
}

impl StdError for PublicationError {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        Some(self.source.as_ref())
    }
}

/// Error returned by the checked streaming provider.
#[derive(Debug)]
pub enum StreamingError {
    /// The typed ODS or XML contract was invalid before or during emission.
    Invalid(Error),
    /// Cancellation or a shared execution budget stopped the stream.
    Execution {
        /// Bytes accepted by the sink before the execution error surfaced.
        written: u64,
        /// The original cancellation/budget error.
        error: ExecutionError,
    },
    /// A finite provider, XML, or package ceiling rejected the next value.
    LimitExceeded {
        /// Stable format-neutral resource name.
        resource: &'static str,
        /// Attempted or observed resource usage.
        observed: u64,
        /// Configured resource ceiling.
        limit: u64,
        /// Bytes accepted by the sink before refusal.
        written: u64,
    },
    /// A checked progress counter could not represent its next value.
    CounterOverflow {
        /// Stable counter name.
        resource: &'static str,
        /// Bytes accepted by the sink before refusal.
        written: u64,
    },
    /// ZIP/publication failure with accepted sink progress.
    Publication(PublicationError),
}

impl fmt::Display for StreamingError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::Execution { written, error } => {
                write!(
                    formatter,
                    "ODS stream stopped after {written} byte(s): {error}"
                )
            },
            Self::LimitExceeded {
                resource,
                observed,
                limit,
                written,
            } => write!(
                formatter,
                "ODS streaming limit for {resource} exceeded: {observed} > {limit} after {written} byte(s)"
            ),
            Self::CounterOverflow { resource, written } => write!(
                formatter,
                "ODS streaming counter {resource} overflowed after {written} byte(s)"
            ),
            Self::Publication(error) => error.fmt(formatter),
        }
    }
}

impl StdError for StreamingError {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        match self {
            Self::Invalid(error) => Some(error),
            Self::Execution { error, .. } => Some(error),
            Self::LimitExceeded { .. } | Self::CounterOverflow { .. } => None,
            Self::Publication(error) => Some(error),
        }
    }
}

fn invalid(message: impl Into<String>) -> StreamingError {
    StreamingError::Invalid(Error::InvalidFormat(message.into()))
}

/// Streams one scalar worksheet to a caller-owned non-seeking sink.
///
/// The iterator is consumed once.  Each row may contain zero through
/// `max_cells_per_row` scalar cells in any order. Number and Boolean values are
/// represented by ODF typed attributes and have no display paragraph. The
/// benchmark harness supplies its own four-cell Number/Text/Boolean/Empty
/// shape; that fixture policy is not imposed on this general writer. The
/// common generated-XML method receives one complete row fragment at a time
/// and retains at most `max_row_xml_bytes` bytes.
///
/// The caller's `ExecutionContext` is checked before metadata publication,
/// before every row and cell, and during scalar escaping. `BudgetedOutput`
/// checks it before every caller-sink write/flush, charges accepted
/// `Resource::OutputBytes`, and performs a final check after archive
/// finalization. The common seam remains responsible for its own bounded
/// fragment and XML checks.
pub fn stream_scalar_rows_to<'a, W, I, R>(
    output: &mut W,
    rows: I,
    context: &ExecutionContext,
    limits: StreamingLimits,
) -> Result<ScalarStreamReport, StreamingError>
where
    W: Write + ?Sized,
    I: IntoIterator<Item = R>,
    R: IntoIterator<Item = StreamingCell<'a>>,
{
    context.check().map_err(|error| execution(error, 0))?;

    // Reserve the common fragment capacity plus the two fixed shell copies
    // that coexist while the common envelope is audited. The reservation is
    // held through entry close and final topology publication; accepted output
    // bytes have their separate OutputBytes reservation in BudgetedOutput.
    let _scratch = reserve_memory(context, limits.required_memory_bytes(), 0)?;

    let envelope = envelope()?;

    // The manifest is a generated member and is longer than content.xml; the
    // transport name budget must admit every fixed package member.
    let member_name_bytes = 32_u64;
    let max_content_bytes = u64::try_from(limits.max_content_xml_bytes)
        .map_err(|_| invalid("content XML limit exceeds u64"))?;
    let max_total_bytes = max_content_bytes
        .checked_add(256 * 1024)
        .ok_or_else(|| invalid("package total byte limit overflow"))?;
    let common_xml_limits = limits.xml_audit.into_common()?;
    let archive_limits = PackageWriterLimits::new(
        3,
        member_name_bytes,
        // This is a provider-side upper bound for the generated manifest
        // bookkeeping.  The common writer still performs its own exact
        // checked candidate admission.
        (64 * 1024).min(limits.max_output_bytes),
    )
    .with_byte_limits(max_content_bytes, max_total_bytes, limits.max_output_bytes)
    .with_compressed_size_limit(limits.max_output_bytes);

    let mut budgeted = BudgetedOutput::new(output, context, limits.max_output_bytes);
    let mut package = PackageWriter::with_writer_and_limits(&mut budgeted, archive_limits);
    let mime_result = package.set_mimetype_streaming(ODS_MIME);
    if let Err(error) = mime_result {
        drop(package);
        return Err(map_package_error(
            error,
            PublicationFailureKind::Preflight,
            &budgeted,
        ));
    }

    let mut producer = ProducerState::new(rows.into_iter(), context, limits);
    let generated_result = package.add_generated_xml(
        CONTENT_PATH,
        CONTENT_MEDIA_TYPE,
        envelope,
        common_xml_limits,
        limits.max_row_xml_bytes,
        |fragment| producer.next_fragment(fragment),
    );
    let generated = match generated_result {
        Ok(generated) => generated,
        Err(error) => {
            drop(package);
            return Err(producer.map_error(error, &budgeted));
        },
    };

    let authored_content_xml_bytes = generated.bytes();
    let rows = producer.rows;
    let cells = producer.cells;

    // `finish_to_writer` consumes the package only after the generated member
    // has closed successfully.  It writes the manifest and central directory
    // through the same checked sink; its progress remains authoritative.
    let finish_result = package.finish_to_writer();
    if let Err(error) = finish_result {
        return Err(producer.map_error_kind(
            error,
            PublicationFailureKind::Finalization,
            &budgeted,
        ));
    }
    if let Err(error) = budgeted.final_check() {
        return Err(execution(error, budgeted.accepted()));
    }

    Ok(ScalarStreamReport {
        rows,
        cells,
        authored_content_xml_bytes,
    })
}

fn reserve_memory(
    context: &ExecutionContext,
    bytes: u64,
    written: u64,
) -> Result<litchi_core::Reservation, StreamingError> {
    context
        .reserve(Resource::Memory, bytes)
        .map_err(|error| execution(error, written))
}

fn execution(error: ExecutionError, written: u64) -> StreamingError {
    StreamingError::Execution { written, error }
}

fn limit_error(limit: LocalLimit, written: u64) -> StreamingError {
    StreamingError::LimitExceeded {
        resource: limit.resource,
        observed: limit.observed,
        limit: limit.limit,
        written,
    }
}

fn envelope() -> Result<GeneratedXmlEnvelope, StreamingError> {
    // The shell has no whitespace-only text nodes and no inherited
    // xml:space state. Rows are independently audited child elements inserted
    // between the table start and end events.
    GeneratedXmlEnvelope::try_new(CONTENT_PREFIX, CONTENT_SUFFIX).map_err(StreamingError::Invalid)
}

#[derive(Clone, Copy)]
struct LocalLimit {
    resource: &'static str,
    observed: u64,
    limit: u64,
}

/// Caller-sink adapter owned by the ODS provider because the common generated
/// XML seam intentionally has no execution-context parameter.  Every write is
/// checked before touching the caller sink, reserves the requested output
/// bytes, commits only the number actually accepted, and retains that count
/// after any failure.
struct BudgetedOutput<'a, W: ?Sized> {
    sink: &'a mut W,
    context: &'a ExecutionContext,
    maximum: u64,
    accepted: u64,
    execution_error: Option<ExecutionError>,
    sink_failed: bool,
    limit_failed: bool,
    limit_observed: Option<u64>,
}

impl<'a, W: ?Sized> BudgetedOutput<'a, W> {
    fn new(sink: &'a mut W, context: &'a ExecutionContext, maximum: u64) -> Self {
        Self {
            sink,
            context,
            maximum,
            accepted: 0,
            execution_error: None,
            sink_failed: false,
            limit_failed: false,
            limit_observed: None,
        }
    }

    fn accepted(&self) -> u64 {
        self.accepted
    }

    fn execution_error(&self) -> Option<ExecutionError> {
        self.execution_error.clone()
    }

    fn sink_failed(&self) -> bool {
        self.sink_failed
    }

    fn limit_failed(&self) -> bool {
        self.limit_failed
    }

    fn limit_observed(&self) -> Option<u64> {
        self.limit_observed
    }

    fn final_check(&mut self) -> Result<(), ExecutionError> {
        match self.context.check() {
            Ok(()) => Ok(()),
            Err(error) => {
                self.execution_error = Some(error.clone());
                Err(error)
            },
        }
    }
}

#[derive(Debug)]
struct BudgetedIoError(ExecutionError);

impl fmt::Display for BudgetedIoError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl StdError for BudgetedIoError {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        Some(&self.0)
    }
}

impl<W: Write + ?Sized> Write for BudgetedOutput<'_, W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        if bytes.is_empty() {
            return Ok(0);
        }
        self.context.check().map_err(|error| {
            self.execution_error = Some(error.clone());
            std::io::Error::other(BudgetedIoError(error))
        })?;

        let amount = u64::try_from(bytes.len())
            .map_err(|_| std::io::Error::other("caller sink write length exceeds u64"))?;
        let next = self.accepted.checked_add(amount).ok_or_else(|| {
            self.limit_failed = true;
            self.limit_observed = Some(u64::MAX);
            std::io::Error::other("accepted output byte counter overflow")
        })?;
        if next > self.maximum {
            self.limit_failed = true;
            self.limit_observed = Some(next);
            return Err(std::io::Error::other(
                "ODS output byte limit exceeded before sink write",
            ));
        }
        let reservation = self
            .context
            .reserve(Resource::OutputBytes, amount)
            .map_err(|error| {
                self.execution_error = Some(error.clone());
                std::io::Error::other(BudgetedIoError(error))
            })?;
        let result = self.sink.write(bytes);
        match result {
            Ok(written) => {
                let written_u64 = u64::try_from(written)
                    .map_err(|_| std::io::Error::other("caller sink write count exceeds u64"))?;
                debug_assert!(written <= bytes.len());
                if written == 0 {
                    self.sink_failed = true;
                }
                if !reservation.commit(written_u64) {
                    self.sink_failed = true;
                    return Err(std::io::Error::other(
                        "output reservation commit exceeded its reservation",
                    ));
                }
                self.accepted = self.accepted.checked_add(written_u64).ok_or_else(|| {
                    self.limit_failed = true;
                    std::io::Error::other("accepted output byte counter overflow")
                })?;
                Ok(written)
            },
            Err(error) => {
                // `Write::write` does not expose a partial count with an
                // error. The reservation is therefore released and the
                // accepted counter remains the last known sink prefix.
                self.sink_failed = true;
                drop(reservation);
                Err(error)
            },
        }
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.context.check().map_err(|error| {
            self.execution_error = Some(error.clone());
            std::io::Error::other(BudgetedIoError(error))
        })?;
        self.sink.flush().inspect_err(|_| {
            self.sink_failed = true;
        })
    }
}

struct ProducerState<'context, I> {
    rows_source: I,
    context: &'context ExecutionContext,
    limits: StreamingLimits,
    rows: usize,
    cells: usize,
    execution_error: Option<ExecutionError>,
    producer_failed: bool,
    limit_error: Option<LocalLimit>,
    counter_error: Option<&'static str>,
}

impl<'context, I> ProducerState<'context, I> {
    fn new(rows_source: I, context: &'context ExecutionContext, limits: StreamingLimits) -> Self {
        Self {
            rows_source,
            context,
            limits,
            rows: 0,
            cells: 0,
            execution_error: None,
            producer_failed: false,
            limit_error: None,
            counter_error: None,
        }
    }

    fn next_fragment<'cell, R>(&mut self, output: &mut dyn Write) -> litchi_core::Result<bool>
    where
        I: Iterator<Item = R>,
        R: IntoIterator<Item = StreamingCell<'cell>>,
    {
        if let Err(error) = self.context.check() {
            self.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }

        if self.rows == self.limits.max_rows {
            // Probe once after the last admitted row.  An extra row is a
            // refusal rather than silent truncation; EOF is the only valid
            // false result.
            if self.rows_source.next().is_some() {
                self.producer_failed = true;
                self.limit_error = Some(LocalLimit {
                    resource: "physical rows",
                    observed: self
                        .rows
                        .checked_add(1)
                        .and_then(|value| u64::try_from(value).ok())
                        .unwrap_or(u64::MAX),
                    limit: u64::try_from(self.limits.max_rows).unwrap_or(u64::MAX),
                });
                return Err(Error::InvalidFormat(
                    "row iterator exceeds max_rows".to_string(),
                ));
            }
            return Ok(false);
        }

        let Some(row) = self.rows_source.next() else {
            return Ok(false);
        };
        let mut row_writer = RowWriter::new(output, self.context, self.limits.max_row_xml_bytes);
        let result = emit_scalar_row(&mut row_writer, row, self.cells, self.limits);
        if let Some(error) = row_writer.execution_error.take() {
            self.execution_error = Some(error.clone());
        }
        if let Some(limit) = row_writer.limit_error.take() {
            self.limit_error = Some(limit);
        }
        if result.is_err() && self.execution_error.is_none() {
            // The callback returned a typed row/limit error.  The common
            // adapter must retain this source while adding accepted output
            // progress; it must not reinterpret it as producer EOF.
            self.producer_failed = true;
        }
        let row_cells = result?;
        self.rows = match self.rows.checked_add(1) {
            Some(value) => value,
            None => {
                self.counter_error = Some("physical rows");
                return Err(Error::InvalidFormat("row counter overflow".to_string()));
            },
        };
        self.cells = match self.cells.checked_add(row_cells) {
            Some(value) => value,
            None => {
                self.counter_error = Some("logical cells");
                return Err(Error::InvalidFormat("cell counter overflow".to_string()));
            },
        };
        Ok(true)
    }

    fn map_error<W: ?Sized>(
        &mut self,
        error: PackageWriterError,
        output: &BudgetedOutput<'_, W>,
    ) -> StreamingError {
        if let Some(limit) = self.limit_error.take() {
            return limit_error(limit, output.accepted());
        }
        if let Some(resource) = self.counter_error.take() {
            return StreamingError::CounterOverflow {
                resource,
                written: output.accepted(),
            };
        }
        if let Some(execution_error) = self.execution_error.take() {
            return execution(execution_error, output.accepted());
        }
        let kind = if self.producer_failed {
            PublicationFailureKind::Producer
        } else {
            PublicationFailureKind::Archive
        };
        map_package_error(error, kind, output)
    }

    fn map_error_kind<W: ?Sized>(
        &mut self,
        error: PackageWriterError,
        kind: PublicationFailureKind,
        output: &BudgetedOutput<'_, W>,
    ) -> StreamingError {
        if let Some(limit) = self.limit_error.take() {
            return limit_error(limit, output.accepted());
        }
        if let Some(resource) = self.counter_error.take() {
            return StreamingError::CounterOverflow {
                resource,
                written: output.accepted(),
            };
        }
        if let Some(execution_error) = self.execution_error.take() {
            return execution(execution_error, output.accepted());
        }
        let kind = if self.producer_failed {
            PublicationFailureKind::Producer
        } else {
            kind
        };
        map_package_error(error, kind, output)
    }
}

struct RowWriter<'a> {
    output: &'a mut dyn Write,
    context: &'a ExecutionContext,
    maximum: usize,
    bytes: usize,
    execution_error: Option<ExecutionError>,
    limit_error: Option<LocalLimit>,
}

impl<'a> RowWriter<'a> {
    fn new(output: &'a mut dyn Write, context: &'a ExecutionContext, maximum: usize) -> Self {
        Self {
            output,
            context,
            maximum,
            bytes: 0,
            execution_error: None,
            limit_error: None,
        }
    }

    fn write_bytes(&mut self, bytes: &[u8]) -> litchi_core::Result<()> {
        if let Err(error) = self.context.check() {
            self.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }
        let next = self
            .bytes
            .checked_add(bytes.len())
            .ok_or_else(|| Error::InvalidFormat("row XML byte counter overflow".to_string()))?;
        if next > self.maximum {
            self.limit_error = Some(LocalLimit {
                resource: "row XML bytes",
                observed: u64::try_from(next).unwrap_or(u64::MAX),
                limit: u64::try_from(self.maximum).unwrap_or(u64::MAX),
            });
            return Err(Error::InvalidFormat(
                "row XML fragment exceeds max_row_xml_bytes".to_string(),
            ));
        }
        let amount = u64::try_from(bytes.len())
            .map_err(|_| Error::InvalidFormat("row XML length exceeds u64".to_string()))?;
        if let Err(error) = self.context.consume(Resource::Work, amount) {
            self.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }
        self.output.write_all(bytes)?;
        self.bytes = next;
        Ok(())
    }

    fn write_ascii(&mut self, text: &str) -> litchi_core::Result<()> {
        self.write_bytes(text.as_bytes())
    }
}

fn emit_scalar_row<'a, R>(
    output: &mut RowWriter<'_>,
    row: R,
    total_cells: usize,
    limits: StreamingLimits,
) -> litchi_core::Result<usize>
where
    R: IntoIterator<Item = StreamingCell<'a>>,
{
    charge_resource(output.context, Resource::Objects, 1, output)?;
    output.write_ascii("<table:table-row>")?;
    let mut cells = row.into_iter();
    let mut row_cells = 0usize;

    loop {
        if let Err(error) = output.context.check() {
            output.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }
        let Some(cell) = cells.next() else {
            break;
        };
        let next_row_cells = row_cells
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("row cell counter overflow".to_string()))?;
        if next_row_cells > limits.max_cells_per_row {
            output.limit_error = Some(LocalLimit {
                resource: "cells per physical row",
                observed: u64::try_from(next_row_cells).unwrap_or(u64::MAX),
                limit: u64::try_from(limits.max_cells_per_row).unwrap_or(u64::MAX),
            });
            return Err(Error::InvalidFormat(
                "row exceeds max_cells_per_row".to_string(),
            ));
        }
        let next_total = total_cells
            .checked_add(next_row_cells)
            .ok_or_else(|| Error::InvalidFormat("cell counter overflow".to_string()))?;
        if next_total > limits.max_cells {
            output.limit_error = Some(LocalLimit {
                resource: "logical cells",
                observed: u64::try_from(next_total).unwrap_or(u64::MAX),
                limit: u64::try_from(limits.max_cells).unwrap_or(u64::MAX),
            });
            return Err(Error::InvalidFormat(
                "scalar stream exceeds max_cells".to_string(),
            ));
        }
        charge_resource(output.context, Resource::Objects, 1, output)?;
        match cell {
            StreamingCell::Number(value) => {
                if !value.is_finite() {
                    return Err(Error::InvalidFormat(
                        "scalar Number must be finite".to_string(),
                    ));
                }
                let lexical = value.to_string();
                output
                    .write_ascii("<table:table-cell office:value-type=\"float\" office:value=\"")?;
                write_attribute_ascii(output, &lexical)?;
                output.write_ascii("\"/>")?;
            },
            StreamingCell::Text(value) => {
                let value = value.as_ref();
                if value.len() > limits.max_text_bytes {
                    output.limit_error = Some(LocalLimit {
                        resource: "text bytes",
                        observed: u64::try_from(value.len()).unwrap_or(u64::MAX),
                        limit: u64::try_from(limits.max_text_bytes).unwrap_or(u64::MAX),
                    });
                    return Err(Error::InvalidFormat(
                        "scalar Text exceeds max_text_bytes".to_string(),
                    ));
                }
                charge_input(output.context, value.len(), output)?;
                validate_xml10_core(value)?;
                output.write_ascii("<table:table-cell office:value-type=\"string\"><text:p")?;
                if requires_xml_space_preserve(value) {
                    output.write_ascii(" xml:space=\"preserve\"")?;
                }
                output.write_ascii(">")?;
                write_text(output, value)?;
                output.write_ascii("</text:p></table:table-cell>")?;
            },
            StreamingCell::Boolean(value) => {
                let lexical = if value { "true" } else { "false" };
                output.write_ascii(
                    "<table:table-cell office:value-type=\"boolean\" office:boolean-value=\"",
                )?;
                output.write_ascii(lexical)?;
                output.write_ascii("\"/>")?;
            },
            StreamingCell::Empty => {
                output.write_ascii("<table:table-cell/>")?;
            },
        }
        row_cells = next_row_cells;
    }
    output.write_ascii("</table:table-row>")?;
    Ok(row_cells)
}

fn charge_resource(
    context: &ExecutionContext,
    resource: Resource,
    amount: u64,
    output: &mut RowWriter<'_>,
) -> litchi_core::Result<()> {
    if let Err(error) = context.consume(resource, amount) {
        output.execution_error = Some(error.clone());
        return Err(Error::Other(error.to_string()));
    }
    Ok(())
}

fn charge_input(
    context: &ExecutionContext,
    bytes: usize,
    output: &mut RowWriter<'_>,
) -> litchi_core::Result<()> {
    let amount = u64::try_from(bytes)
        .map_err(|_| Error::InvalidFormat("text length exceeds u64".to_string()))?;
    charge_resource(context, Resource::InputBytes, amount, output)
}

fn write_attribute_ascii(output: &mut RowWriter<'_>, value: &str) -> litchi_core::Result<()> {
    // Number lexical values contain only ASCII under f64::to_string.  The
    // helper still rejects unexpected control characters so this remains a
    // checked attribute boundary if the lexical producer changes.
    for byte in value.bytes() {
        if !(byte.is_ascii_graphic() || byte == b' ') || byte == b'"' || byte == b'&' {
            return Err(Error::InvalidFormat(
                "invalid character in generated XML attribute".to_string(),
            ));
        }
    }
    output.write_bytes(value.as_bytes())
}

fn write_text(output: &mut RowWriter<'_>, value: &str) -> litchi_core::Result<()> {
    // Encode one scalar at a time.  This avoids a second text-sized retained
    // buffer when a text value contains no entities; the common FragmentBuffer
    // remains the only row XML window.
    for character in value.chars() {
        match character {
            '&' => output.write_ascii("&amp;")?,
            '<' => output.write_ascii("&lt;")?,
            '>' => output.write_ascii("&gt;")?,
            '"' => output.write_ascii("&quot;")?,
            '\'' => output.write_ascii("&apos;")?,
            '\r' => output.write_ascii("&#13;")?,
            '\n' => output.write_ascii("&#10;")?,
            '\t' => output.write_ascii("&#9;")?,
            character => {
                let mut encoded = [0u8; 4];
                let encoded = character.encode_utf8(&mut encoded);
                output.write_bytes(encoded.as_bytes())?;
            },
        }
    }
    Ok(())
}

fn requires_xml_space_preserve(value: &str) -> bool {
    value
        .bytes()
        .any(|byte| matches!(byte, b' ' | b'\t' | b'\n' | b'\r'))
}

fn validate_xml10_core(value: &str) -> litchi_core::Result<()> {
    if value.chars().all(is_xml10_character) {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "text contains an XML 1.0-incompatible character".to_string(),
        ))
    }
}

const fn is_xml10_character(character: char) -> bool {
    matches!(
        character as u32,
        0x09 | 0x0A | 0x0D | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
}

fn map_package_error<W: ?Sized>(
    error: PackageWriterError,
    kind: PublicationFailureKind,
    output: &BudgetedOutput<'_, W>,
) -> StreamingError {
    if let Some(execution_error) = output.execution_error() {
        return execution(execution_error, output.accepted());
    }
    let written = output.accepted().max(error.written().unwrap_or(0));
    if output.limit_failed() {
        return StreamingError::LimitExceeded {
            resource: "output bytes",
            observed: output.limit_observed().unwrap_or(u64::MAX),
            limit: output.maximum,
            written,
        };
    }
    if let Some(limit) = error.limit() {
        let resource = match limit.resource() {
            litchi_odf_common::core::PackageWriterLimitResource::FileCount => "file count",
            litchi_odf_common::core::PackageWriterLimitResource::MemberNameBytes => {
                "member name bytes"
            },
            litchi_odf_common::core::PackageWriterLimitResource::MetadataBytes => "metadata bytes",
            litchi_odf_common::core::PackageWriterLimitResource::CompressedSize => {
                "compressed member bytes"
            },
            litchi_odf_common::core::PackageWriterLimitResource::EntrySize => {
                "uncompressed member bytes"
            },
            litchi_odf_common::core::PackageWriterLimitResource::TotalSize => {
                "total uncompressed bytes"
            },
            litchi_odf_common::core::PackageWriterLimitResource::OutputBytes => "output bytes",
            litchi_odf_common::core::PackageWriterLimitResource::Other => "package resource",
        };
        return StreamingError::LimitExceeded {
            resource,
            observed: limit.actual(),
            limit: limit.maximum(),
            written,
        };
    }
    let kind = if error.limit().is_some() || output.limit_failed() {
        PublicationFailureKind::Limit
    } else if output.sink_failed() {
        PublicationFailureKind::Sink
    } else {
        kind
    };
    StreamingError::Publication(PublicationError {
        written,
        kind,
        source: Box::new(error),
    })
}
