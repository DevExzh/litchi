//! Bounded, forward-only authoring of plain DOCX paragraphs and runs.
//!
//! [`StreamingDocumentWriter`] is deliberately narrower than the regular
//! [`crate::Package`] and [`crate::writer`] APIs.  It creates a fresh
//! transitional WordprocessingML package over a sequential sink.  The writer
//! owns no document model and cannot revisit a paragraph or run after it has
//! been flushed.

use std::fmt;
use std::io::{self, Write};
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicU64, Ordering},
};

use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource};
use litchi_opc::phys_pkg::{PartWriter, PhysPkgWriter};
use litchi_opc::{OpcError, PackURI};

const FIXED_SCRATCH_BYTES: u64 = 64;
const TEXT_SCRATCH_BYTES: usize = 64;
const MIN_OUTPUT_BYTES: u64 = 22;
// `PhysPkgWriter::with_writer` intentionally uses these finite ZIP transport
// defaults.  Keep the semantic API honest until OPC exposes a configurable
// physical-writer constructor through this crate.
const OPC_MAX_ENTRY_BYTES: u64 = 512 * 1024 * 1024;
const OPC_MAX_OUTPUT_BYTES: u64 = 512 * 1024 * 1024;
const FIXED_OBJECT_CHARGE: u64 = 5;
const FIXED_WORK_CHARGE: u64 = 2;
const TEXT_SCAN_CHECK_INTERVAL: u64 = 64;

const CONTENT_TYPES_URI: &str = "/[Content_Types].xml";
const PACKAGE_RELS_URI: &str = "/_rels/.rels";
const DOCUMENT_URI: &str = "/word/document.xml";

const CONTENT_TYPES_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
    r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
    r#"<Default Extension="xml" ContentType="application/xml"/>"#,
    r#"<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>"#,
    r#"</Types>"#,
)
.as_bytes();

const PACKAGE_RELS_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>"#,
    r#"</Relationships>"#,
)
.as_bytes();

const DOCUMENT_PREFIX: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#,
)
.as_bytes();
const DOCUMENT_SUFFIX: &[u8] = b"<w:sectPr/></w:body></w:document>";
const PARAGRAPH_PREFIX: &[u8] = b"<w:p>";
const PARAGRAPH_SUFFIX: &[u8] = b"</w:p>";
const RUN_PREFIX: &[u8] = b"<w:r><w:t xml:space=\"preserve\">";
const RUN_SUFFIX: &[u8] = b"</w:t></w:r>";

/// Finite semantic limits for [`StreamingDocumentWriter`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StreamingDocumentLimits {
    /// Maximum UTF-8 input bytes accepted by all runs.
    pub max_input_bytes: u64,
    /// Maximum bytes accepted by the complete output sink.
    pub max_output_bytes: u64,
    /// Maximum number of paragraphs.
    pub max_paragraphs: u64,
    /// Maximum number of runs.
    pub max_runs: u64,
    /// Maximum number of runs in one paragraph.
    pub max_runs_per_paragraph: u64,
    /// Maximum UTF-8 bytes accepted in one run.
    pub max_run_text_bytes: u64,
    /// Maximum uncompressed XML bytes in one paragraph.
    pub max_paragraph_xml_bytes: u64,
    /// Maximum uncompressed bytes in `/word/document.xml`.
    pub max_document_xml_bytes: u64,
    /// Minimum scratch capability accepted by this writer.
    pub max_scratch_bytes: u64,
}

impl StreamingDocumentLimits {
    /// Creates an explicit finite limit set.
    #[must_use]
    pub const fn new(
        max_input_bytes: u64,
        max_output_bytes: u64,
        max_paragraphs: u64,
        max_runs: u64,
        max_runs_per_paragraph: u64,
        max_run_text_bytes: u64,
        max_paragraph_xml_bytes: u64,
        max_document_xml_bytes: u64,
        max_scratch_bytes: u64,
    ) -> Self {
        Self {
            max_input_bytes,
            max_output_bytes,
            max_paragraphs,
            max_runs,
            max_runs_per_paragraph,
            max_run_text_bytes,
            max_paragraph_xml_bytes,
            max_document_xml_bytes,
            max_scratch_bytes,
        }
    }
}

impl Default for StreamingDocumentLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: 256 * 1024 * 1024,
            max_output_bytes: 256 * 1024 * 1024,
            max_paragraphs: 1_000_000,
            max_runs: 4_000_000,
            max_runs_per_paragraph: 4_000_000,
            max_run_text_bytes: 16 * 1024 * 1024,
            max_paragraph_xml_bytes: 64 * 1024 * 1024,
            max_document_xml_bytes: 256 * 1024 * 1024,
            max_scratch_bytes: FIXED_SCRATCH_BYTES,
        }
    }
}

/// Shared source-chain handle carried by [`StreamingDocumentError`].
pub type StreamingDocumentErrorSource = Arc<dyn std::error::Error + Send + Sync>;

fn error_source<E>(error: E) -> StreamingDocumentErrorSource
where
    E: std::error::Error + Send + Sync + 'static,
{
    Arc::new(error)
}

fn as_error_source<'a>(
    error: &'a (dyn std::error::Error + Send + Sync + 'static),
) -> &'a (dyn std::error::Error + 'static) {
    error
}

#[derive(Debug)]
struct ErrorChainLink {
    current: StreamingDocumentErrorSource,
    next: StreamingDocumentErrorSource,
}

impl fmt::Display for ErrorChainLink {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.current.fmt(formatter)
    }
}

impl std::error::Error for ErrorChainLink {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(as_error_source(&self.next))
    }
}

fn chained_error_source(
    current: impl std::error::Error + Send + Sync + 'static,
    next: Option<impl std::error::Error + Send + Sync + 'static>,
) -> StreamingDocumentErrorSource {
    let current = error_source(current);
    match next {
        Some(next) => Arc::new(ErrorChainLink {
            current,
            next: error_source(next),
        }),
        None => current,
    }
}

/// Failure from bounded forward-only DOCX authoring.
#[derive(Debug)]
#[non_exhaustive]
pub enum StreamingDocumentError {
    /// The sink failed after accepting the reported output bytes.
    Io {
        /// Underlying I/O category.
        kind: io::ErrorKind,
        /// Stable diagnostic captured from the sink.
        message: String,
        /// Output bytes accepted before the failure.
        written: u64,
        /// Preserved sink error chain, when available.
        source: Option<StreamingDocumentErrorSource>,
    },
    /// Caller input or a state transition violated the narrow authoring
    /// contract.
    InvalidInput {
        /// Human-readable reason.
        message: String,
        /// Output bytes accepted before the failure.
        written: u64,
    },
    /// A local or hierarchical finite resource budget was exhausted.
    LimitExceeded {
        /// Stable resource name.
        resource: &'static str,
        /// Observed value reported by the budget.
        observed: u64,
        /// Configured maximum.
        limit: u64,
        /// Output bytes accepted before the failure.
        written: u64,
    },
    /// The supplied execution context was cancelled.
    Cancelled {
        /// Output bytes accepted before cancellation was observed.
        written: u64,
    },
    /// The OPC physical writer rejected or could not finalize the package.
    Opc {
        /// Stable diagnostic captured from the physical writer.
        message: String,
        /// Output bytes accepted before the failure.
        written: u64,
        /// Preserved OPC error chain, when available.
        source: Option<StreamingDocumentErrorSource>,
    },
    /// The physical package became incomplete after output had been accepted.
    IncompleteOutput {
        /// Stable diagnostic captured from the physical writer.
        message: String,
        /// Output bytes accepted before the incomplete publication.
        written: u64,
        /// Preserved physical-publication error chain, when available.
        source: Option<StreamingDocumentErrorSource>,
    },
    /// A bounded progress counter could not represent the next value.
    CounterOverflow {
        /// Counter whose next value was not representable.
        resource: &'static str,
        /// Output bytes accepted before the failure.
        written: u64,
    },
}

impl StreamingDocumentError {
    /// Output bytes accepted before this failure.
    #[must_use]
    pub const fn written(&self) -> u64 {
        match self {
            Self::Io { written, .. }
            | Self::InvalidInput { written, .. }
            | Self::LimitExceeded { written, .. }
            | Self::Cancelled { written }
            | Self::Opc { written, .. }
            | Self::IncompleteOutput { written, .. }
            | Self::CounterOverflow { written, .. } => *written,
        }
    }
}

impl fmt::Display for StreamingDocumentError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Io {
                kind,
                message,
                written,
                ..
            } => write!(
                formatter,
                "streaming DOCX sink failed ({kind}) after {written} output bytes: {message}"
            ),
            Self::InvalidInput { message, written } => write!(
                formatter,
                "invalid streaming DOCX authoring input after {written} output bytes: {message}"
            ),
            Self::LimitExceeded {
                resource,
                observed,
                limit,
                written,
            } => write!(
                formatter,
                "streaming DOCX limit exceeded for {resource}: observed {observed}, limit {limit}, output {written}"
            ),
            Self::Cancelled { written } => write!(
                formatter,
                "streaming DOCX authoring cancelled after {written} output bytes"
            ),
            Self::Opc {
                message, written, ..
            } => write!(
                formatter,
                "streaming DOCX OPC publication failed after {written} output bytes: {message}"
            ),
            Self::IncompleteOutput {
                message, written, ..
            } => write!(
                formatter,
                "streaming DOCX output became incomplete after {written} output bytes: {message}"
            ),
            Self::CounterOverflow { resource, written } => write!(
                formatter,
                "streaming DOCX progress counter overflowed for {resource} after {written} output bytes"
            ),
        }
    }
}

impl std::error::Error for StreamingDocumentError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        let source = match self {
            Self::Io { source, .. }
            | Self::Opc { source, .. }
            | Self::IncompleteOutput { source, .. } => source.as_deref(),
            Self::InvalidInput { .. }
            | Self::LimitExceeded { .. }
            | Self::Cancelled { .. }
            | Self::CounterOverflow { .. } => None,
        };
        source.map(as_error_source)
    }
}

#[derive(Debug, Clone)]
enum PoisonReason {
    Io {
        kind: io::ErrorKind,
        message: String,
        source: Option<StreamingDocumentErrorSource>,
    },
    InvalidInput(String),
    Limit {
        resource: &'static str,
        observed: u64,
        limit: u64,
    },
    Cancelled,
    Opc {
        message: String,
        source: Option<StreamingDocumentErrorSource>,
    },
    IncompleteOutput {
        message: String,
        source: Option<StreamingDocumentErrorSource>,
    },
    CounterOverflow(&'static str),
}

impl PoisonReason {
    fn error(&self, written: u64) -> StreamingDocumentError {
        match self {
            Self::Io {
                kind,
                message,
                source,
            } => StreamingDocumentError::Io {
                kind: *kind,
                message: message.clone(),
                written,
                source: source.clone(),
            },
            Self::InvalidInput(message) => StreamingDocumentError::InvalidInput {
                message: message.clone(),
                written,
            },
            Self::Limit {
                resource,
                observed,
                limit,
            } => StreamingDocumentError::LimitExceeded {
                resource,
                observed: *observed,
                limit: *limit,
                written,
            },
            Self::Cancelled => StreamingDocumentError::Cancelled { written },
            Self::Opc { message, source } => StreamingDocumentError::Opc {
                message: message.clone(),
                written,
                source: source.clone(),
            },
            Self::IncompleteOutput { message, source } => {
                StreamingDocumentError::IncompleteOutput {
                    message: message.clone(),
                    written,
                    source: source.clone(),
                }
            },
            Self::CounterOverflow(resource) => {
                StreamingDocumentError::CounterOverflow { resource, written }
            },
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Phase {
    Ready,
    Paragraph {
        run_open: bool,
        run_count: u64,
        run_text_bytes: u64,
        paragraph_xml_bytes: u64,
    },
}

/// A sequential, bounded writer for a fresh plain-text DOCX document.
///
/// The writer owns the caller's non-seek sink.  It emits the two static OPC
/// manifests first and then streams `/word/document.xml` as the final member.
/// Only plain `w:t` runs are accepted.  Paragraphs and runs are consumed in
/// order and cannot be revisited after their closing method succeeds.
///
/// The execution context charges one writer object, one object per paragraph,
/// one object per run, package setup/finalization work, input bytes, output
/// bytes, and a fixed XML-escaping scratch capability.  Dropping the writer
/// before [`Self::finish`] abandons the package and can leave the owned sink
/// with an intentionally incomplete archive.
///
/// The semantic counters and scratch reservation are bounded here, but this
/// API does not promise constant total RSS: the caller's sink and the
/// underlying ZIP/Deflate implementation retain transport state outside this
/// writer's accounting surface.
pub struct StreamingDocumentWriter<W: Write> {
    part: Option<PartWriter<BudgetedOutput<W>>>,
    context: ExecutionContext,
    limits: StreamingDocumentLimits,
    phase: Phase,
    paragraphs: u64,
    runs: u64,
    input_bytes: u64,
    document_xml_bytes: u64,
    output_counter: Arc<AtomicU64>,
    execution_state: Arc<ExecutionFailureState>,
    _scratch: Reservation,
    poison: Option<PoisonReason>,
}

impl<W: Write> StreamingDocumentWriter<W> {
    /// Creates a writer and immediately publishes the static package members
    /// and the document XML prefix.
    pub fn new(
        writer: W,
        context: ExecutionContext,
        limits: StreamingDocumentLimits,
    ) -> Result<Self, StreamingDocumentError> {
        validate_limits(limits)?;
        context.check().map_err(|error| map_execution(error, 0))?;
        let scratch = context
            .reserve(Resource::Memory, FIXED_SCRATCH_BYTES)
            .map_err(|error| map_execution(error, 0))?;
        // Charge every fixed object/work unit before the first ZIP header is
        // published.  Dynamic paragraph/run charges remain at their state
        // transitions, while package finalization is reserved up front.
        context
            .consume(Resource::Objects, FIXED_OBJECT_CHARGE)
            .and_then(|_| context.consume(Resource::Work, FIXED_WORK_CHARGE))
            .map_err(|error| map_execution(error, 0))?;

        let output_counter = Arc::new(AtomicU64::new(0));
        let execution_state = Arc::new(ExecutionFailureState::default());
        let document_prefix_bytes = usize_to_u64(DOCUMENT_PREFIX.len(), "document XML bytes")
            .map_err(|resource| StreamingDocumentError::CounterOverflow {
                resource,
                written: output_counter.load(Ordering::Acquire),
            })?;
        let output = BudgetedOutput::new(
            writer,
            context.clone(),
            limits.max_output_bytes,
            Arc::clone(&output_counter),
            Arc::clone(&execution_state),
        );
        let mut physical = PhysPkgWriter::with_writer(output);
        for (name, bytes) in [
            (CONTENT_TYPES_URI, CONTENT_TYPES_XML),
            (PACKAGE_RELS_URI, PACKAGE_RELS_XML),
        ] {
            let uri = PackURI::new(name).map_err(|error| {
                invalid_with_written(error.to_string(), output_counter.load(Ordering::Acquire))
            })?;
            let mut member = physical.start_part(&uri).map_err(|error| {
                map_opc_with_state(
                    error,
                    output_counter.load(Ordering::Acquire),
                    &execution_state,
                )
            })?;
            if let Err(write_error) = member.write_all(bytes) {
                let finish_error = member.finish().err();
                return Err(map_member_write_error(
                    write_error,
                    finish_error,
                    output_counter.load(Ordering::Acquire),
                    &execution_state,
                ));
            }
            physical = member.finish().map_err(|error| {
                map_opc_with_state(
                    error,
                    output_counter.load(Ordering::Acquire),
                    &execution_state,
                )
            })?;
        }

        let document_uri = PackURI::new(DOCUMENT_URI).map_err(|error| {
            invalid_with_written(error.to_string(), output_counter.load(Ordering::Acquire))
        })?;
        let mut part = physical.start_part(&document_uri).map_err(|error| {
            map_opc_with_state(
                error,
                output_counter.load(Ordering::Acquire),
                &execution_state,
            )
        })?;
        if let Err(write_error) = part.write_all(DOCUMENT_PREFIX) {
            let finish_error = part.finish().err();
            return Err(map_member_write_error(
                write_error,
                finish_error,
                output_counter.load(Ordering::Acquire),
                &execution_state,
            ));
        }
        Ok(Self {
            part: Some(part),
            context,
            limits,
            phase: Phase::Ready,
            paragraphs: 0,
            runs: 0,
            input_bytes: 0,
            document_xml_bytes: document_prefix_bytes,
            output_counter,
            execution_state,
            _scratch: scratch,
            poison: None,
        })
    }

    /// Number of paragraphs successfully started.
    #[must_use]
    pub const fn paragraph_count(&self) -> u64 {
        self.paragraphs
    }

    /// Number of runs successfully started.
    #[must_use]
    pub const fn run_count(&self) -> u64 {
        self.runs
    }

    /// UTF-8 input bytes successfully committed to completed text calls.
    #[must_use]
    pub const fn input_bytes(&self) -> u64 {
        self.input_bytes
    }

    /// Uncompressed `/word/document.xml` bytes accepted so far.
    #[must_use]
    pub const fn document_xml_bytes(&self) -> u64 {
        self.document_xml_bytes
    }

    /// Number of bytes accepted by the caller's sink.
    #[must_use]
    pub fn output_bytes(&self) -> u64 {
        self.output_counter.load(Ordering::Acquire)
    }

    /// Whether a sink, cancellation, limit, or finalization failure poisoned
    /// this writer.
    #[must_use]
    pub const fn is_poisoned(&self) -> bool {
        self.poison.is_some()
    }

    /// Start the next logical paragraph.
    pub fn start_paragraph(&mut self) -> Result<(), StreamingDocumentError> {
        self.ensure_usable()?;
        if !matches!(self.phase, Phase::Ready) {
            return Err(self.state_error("a paragraph is already open"));
        }
        self.context
            .check()
            .map_err(|error| self.fail_execution(error))?;
        let observed = checked_increment(self.paragraphs, "paragraphs")
            .map_err(|resource| self.fail_overflow(resource))?;
        if observed > self.limits.max_paragraphs {
            return Err(self.fail_limit("paragraphs", observed, self.limits.max_paragraphs));
        }
        let paragraph_bytes = usize_to_u64(PARAGRAPH_PREFIX.len(), "paragraph XML bytes")
            .map_err(|resource| self.fail_overflow(resource))?;
        self.preflight_document_and_paragraph(paragraph_bytes, paragraph_bytes)?;
        self.context
            .consume(Resource::Objects, 1)
            .and_then(|_| self.context.consume(Resource::Work, 1))
            .map_err(|error| self.fail_execution(error))?;
        self.emit_document(PARAGRAPH_PREFIX, paragraph_bytes, Some(paragraph_bytes))?;
        self.paragraphs = observed;
        self.phase = Phase::Paragraph {
            run_open: false,
            run_count: 0,
            run_text_bytes: 0,
            paragraph_xml_bytes: paragraph_bytes,
        };
        Ok(())
    }

    /// Start a plain text run in the current paragraph.
    pub fn start_run(&mut self) -> Result<(), StreamingDocumentError> {
        self.ensure_usable()?;
        let Phase::Paragraph {
            run_open,
            run_count,
            paragraph_xml_bytes,
            ..
        } = self.phase
        else {
            return Err(self.state_error("a paragraph must be open before a run"));
        };
        if run_open {
            return Err(self.state_error("a run is already open"));
        }
        self.context
            .check()
            .map_err(|error| self.fail_execution(error))?;
        let observed_runs = checked_increment(self.runs, "runs")
            .map_err(|resource| self.fail_overflow(resource))?;
        if observed_runs > self.limits.max_runs {
            return Err(self.fail_limit("runs", observed_runs, self.limits.max_runs));
        }
        let observed_paragraph_runs = checked_increment(run_count, "runs per paragraph")
            .map_err(|resource| self.fail_overflow(resource))?;
        if observed_paragraph_runs > self.limits.max_runs_per_paragraph {
            return Err(self.fail_limit(
                "runs per paragraph",
                observed_paragraph_runs,
                self.limits.max_runs_per_paragraph,
            ));
        }
        let run_bytes = usize_to_u64(RUN_PREFIX.len(), "paragraph XML bytes")
            .map_err(|resource| self.fail_overflow(resource))?;
        let next_paragraph_bytes =
            checked_add(paragraph_xml_bytes, run_bytes, "paragraph XML bytes")
                .map_err(|resource| self.fail_overflow(resource))?;
        self.preflight_document_and_paragraph(run_bytes, next_paragraph_bytes)?;
        self.context
            .consume(Resource::Objects, 1)
            .and_then(|_| self.context.consume(Resource::Work, 1))
            .map_err(|error| self.fail_execution(error))?;
        self.emit_document(RUN_PREFIX, run_bytes, Some(run_bytes))?;
        self.runs = observed_runs;
        self.phase = Phase::Paragraph {
            run_open: true,
            run_count: observed_paragraph_runs,
            run_text_bytes: 0,
            paragraph_xml_bytes: next_paragraph_bytes,
        };
        Ok(())
    }

    /// Write one UTF-8 plain-text fragment to the current run.
    ///
    /// The complete fragment is validated and sized before any of its XML is
    /// emitted.  Tabs, line feeds, and carriage returns are rejected because
    /// they require structural `w:tab`, `w:br`, or `w:cr` elements.
    pub fn write_text(&mut self, text: &str) -> Result<(), StreamingDocumentError> {
        self.ensure_usable()?;
        let Phase::Paragraph {
            run_open: true,
            run_text_bytes,
            paragraph_xml_bytes,
            ..
        } = self.phase
        else {
            return Err(self.state_error("text may only be written while a run is open"));
        };
        self.context
            .check()
            .map_err(|error| self.fail_execution(error))?;
        let input_bytes = usize_to_u64(text.len(), "input bytes")
            .map_err(|resource| self.fail_overflow(resource))?;
        let next_run_bytes = checked_add(run_text_bytes, input_bytes, "run text bytes")
            .map_err(|resource| self.fail_overflow(resource))?;
        if next_run_bytes > self.limits.max_run_text_bytes {
            return Err(self.fail_limit(
                "run text bytes",
                next_run_bytes,
                self.limits.max_run_text_bytes,
            ));
        }
        let next_input = checked_add(self.input_bytes, input_bytes, "input bytes")
            .map_err(|resource| self.fail_overflow(resource))?;
        if next_input > self.limits.max_input_bytes {
            return Err(self.fail_limit("input bytes", next_input, self.limits.max_input_bytes));
        }
        let (encoded_bytes, _text_chars) = self.encoded_text_size(text)?;
        let next_paragraph_bytes =
            checked_add(paragraph_xml_bytes, encoded_bytes, "paragraph XML bytes")
                .map_err(|resource| self.fail_overflow(resource))?;
        self.preflight_document_and_paragraph(encoded_bytes, next_paragraph_bytes)?;
        let input_reservation = self
            .context
            .reserve(Resource::InputBytes, input_bytes)
            .map_err(|error| self.fail_execution(error))?;
        self.emit_escaped_text(text)?;
        if !input_reservation.commit(input_bytes) {
            return Err(self.fail_invalid("input reservation committed more bytes than reserved"));
        }
        self.input_bytes = next_input;
        self.phase = match self.phase {
            Phase::Paragraph {
                run_open,
                run_count,
                paragraph_xml_bytes,
                ..
            } => Phase::Paragraph {
                run_open,
                run_count,
                run_text_bytes: next_run_bytes,
                // `emit_document` owns document/paragraph XML progress.  The
                // escaped chunks may cross several scratch boundaries, so do
                // not add `encoded_bytes` a second time here.
                paragraph_xml_bytes,
            },
            Phase::Ready => Phase::Ready,
        };
        Ok(())
    }

    /// Finish the current run.
    pub fn finish_run(&mut self) -> Result<(), StreamingDocumentError> {
        self.ensure_usable()?;
        let Phase::Paragraph {
            run_open: true,
            run_count,
            run_text_bytes,
            paragraph_xml_bytes,
        } = self.phase
        else {
            return Err(self.state_error("no run is open"));
        };
        let suffix_bytes = usize_to_u64(RUN_SUFFIX.len(), "paragraph XML bytes")
            .map_err(|resource| self.fail_overflow(resource))?;
        let next_paragraph_bytes =
            checked_add(paragraph_xml_bytes, suffix_bytes, "paragraph XML bytes")
                .map_err(|resource| self.fail_overflow(resource))?;
        self.context
            .check()
            .map_err(|error| self.fail_execution(error))?;
        self.preflight_document_and_paragraph(suffix_bytes, next_paragraph_bytes)?;
        self.context
            .consume(Resource::Work, 1)
            .map_err(|error| self.fail_execution(error))?;
        self.emit_document(RUN_SUFFIX, suffix_bytes, Some(suffix_bytes))?;
        self.phase = Phase::Paragraph {
            run_open: false,
            run_count,
            run_text_bytes,
            paragraph_xml_bytes: next_paragraph_bytes,
        };
        Ok(())
    }

    /// Finish the current paragraph.
    pub fn finish_paragraph(&mut self) -> Result<(), StreamingDocumentError> {
        self.ensure_usable()?;
        let Phase::Paragraph {
            run_open: false,
            paragraph_xml_bytes,
            ..
        } = self.phase
        else {
            return Err(self.state_error("finish the open run before the paragraph"));
        };
        let suffix_bytes = usize_to_u64(PARAGRAPH_SUFFIX.len(), "paragraph XML bytes")
            .map_err(|resource| self.fail_overflow(resource))?;
        let next_paragraph_bytes =
            checked_add(paragraph_xml_bytes, suffix_bytes, "paragraph XML bytes")
                .map_err(|resource| self.fail_overflow(resource))?;
        self.context
            .check()
            .map_err(|error| self.fail_execution(error))?;
        self.preflight_document_and_paragraph(suffix_bytes, next_paragraph_bytes)?;
        self.context
            .consume(Resource::Work, 1)
            .map_err(|error| self.fail_execution(error))?;
        self.emit_document(PARAGRAPH_SUFFIX, suffix_bytes, Some(suffix_bytes))?;
        self.phase = Phase::Ready;
        Ok(())
    }

    /// Finalize the document XML, ZIP central directory, and caller-owned sink.
    pub fn finish(mut self) -> Result<W, StreamingDocumentError> {
        self.ensure_usable()?;
        if !matches!(self.phase, Phase::Ready) {
            return Err(self.state_error("finish the current paragraph first"));
        }
        self.context
            .check()
            .map_err(|error| self.fail_execution(error))?;
        let suffix_bytes = usize_to_u64(DOCUMENT_SUFFIX.len(), "document XML bytes")
            .map_err(|resource| self.fail_overflow(resource))?;
        let Some(observed) = self.document_xml_bytes.checked_add(suffix_bytes) else {
            return Err(self.fail_overflow("document XML bytes"));
        };
        if observed > self.limits.max_document_xml_bytes {
            return Err(self.fail_limit(
                "document XML bytes",
                observed,
                self.limits.max_document_xml_bytes,
            ));
        }
        self.emit_document(DOCUMENT_SUFFIX, suffix_bytes, None)?;
        let Some(part) = self.part.take() else {
            return Err(self.fail_invalid("streaming DOCX document part is unavailable"));
        };
        let physical = match part.finish() {
            Ok(physical) => physical,
            Err(error) => {
                let mapped = map_opc_with_state(error, self.output_bytes(), &self.execution_state);
                self.poison = Some(poison_from_error(&mapped));
                return Err(mapped);
            },
        };
        let output = match physical.finish_into_inner() {
            Ok(output) => output,
            Err(error) => {
                let mapped = map_opc_with_state(error, self.output_bytes(), &self.execution_state);
                self.poison = Some(poison_from_error(&mapped));
                return Err(mapped);
            },
        };
        let mut output = output;
        if let Err(error) = output.flush() {
            let mapped =
                map_member_write_error(error, None, self.output_bytes(), &self.execution_state);
            self.poison = Some(poison_from_error(&mapped));
            return Err(mapped);
        }
        output.into_inner().map_err(|error| {
            let mapped =
                map_member_write_error(error, None, self.output_bytes(), &self.execution_state);
            self.poison = Some(poison_from_error(&mapped));
            mapped
        })
    }

    fn ensure_usable(&self) -> Result<(), StreamingDocumentError> {
        self.poison
            .as_ref()
            .map_or(Ok(()), |reason| Err(reason.error(self.output_bytes())))
    }

    fn state_error(&self, message: &str) -> StreamingDocumentError {
        StreamingDocumentError::InvalidInput {
            message: message.to_owned(),
            written: self.output_bytes(),
        }
    }

    fn fail_invalid(&mut self, message: impl Into<String>) -> StreamingDocumentError {
        self.poison = Some(PoisonReason::InvalidInput(message.into()));
        self.poison.as_ref().map_or_else(
            || self.state_error("streaming DOCX input failed"),
            |reason| reason.error(self.output_bytes()),
        )
    }

    fn fail_limit(
        &mut self,
        resource: &'static str,
        observed: u64,
        limit: u64,
    ) -> StreamingDocumentError {
        self.poison = Some(PoisonReason::Limit {
            resource,
            observed,
            limit,
        });
        self.poison.as_ref().map_or_else(
            || self.state_error("streaming DOCX limit failed"),
            |reason| reason.error(self.output_bytes()),
        )
    }

    fn fail_overflow(&mut self, resource: &'static str) -> StreamingDocumentError {
        self.poison = Some(PoisonReason::CounterOverflow(resource));
        self.poison.as_ref().map_or_else(
            || self.state_error("streaming DOCX progress counter overflowed"),
            |reason| reason.error(self.output_bytes()),
        )
    }

    fn fail_execution(&mut self, error: ExecutionError) -> StreamingDocumentError {
        let reason = poison_from_execution(error);
        self.poison = Some(reason);
        self.poison.as_ref().map_or_else(
            || self.state_error("streaming DOCX execution failed"),
            |reason| reason.error(self.output_bytes()),
        )
    }

    fn preflight_document_and_paragraph(
        &mut self,
        additional_document_bytes: u64,
        paragraph_bytes: u64,
    ) -> Result<(), StreamingDocumentError> {
        let document_observed = match self
            .document_xml_bytes
            .checked_add(additional_document_bytes)
        {
            Some(value) => value,
            None => return Err(self.fail_overflow("document XML bytes")),
        };
        if document_observed > self.limits.max_document_xml_bytes {
            return Err(self.fail_limit(
                "document XML bytes",
                document_observed,
                self.limits.max_document_xml_bytes,
            ));
        }
        if paragraph_bytes > self.limits.max_paragraph_xml_bytes {
            return Err(self.fail_limit(
                "paragraph XML bytes",
                paragraph_bytes,
                self.limits.max_paragraph_xml_bytes,
            ));
        }
        Ok(())
    }

    fn emit_document(
        &mut self,
        bytes: &[u8],
        document_bytes: u64,
        paragraph_bytes: Option<u64>,
    ) -> Result<(), StreamingDocumentError> {
        if bytes.is_empty() {
            return Ok(());
        }
        let (result, accepted) = match self.part.as_mut() {
            Some(part) => {
                let before = part.uncompressed_bytes();
                let result = part.write_all(bytes);
                let after = part.uncompressed_bytes();
                let accepted = after.checked_sub(before).unwrap_or(0);
                (result, accepted)
            },
            None => {
                return Err(self.fail_invalid("streaming DOCX document part is unavailable"));
            },
        };
        if let Err(error) = result {
            // A ZIP/Deflate sink can accept a prefix before reporting an I/O
            // error.  Preserve that semantic prefix instead of reporting the
            // whole chunk as if no document XML had been accepted.
            self.record_document_progress(accepted, paragraph_bytes.map(|_| accepted))?;
            return Err(self.fail_io(error));
        }
        self.record_document_progress(document_bytes, paragraph_bytes)
    }

    fn record_document_progress(
        &mut self,
        document_bytes: u64,
        paragraph_bytes: Option<u64>,
    ) -> Result<(), StreamingDocumentError> {
        self.document_xml_bytes = match self.document_xml_bytes.checked_add(document_bytes) {
            Some(value) => value,
            None => return Err(self.fail_overflow("document XML bytes")),
        };
        if let Some(paragraph_bytes) = paragraph_bytes
            && let Phase::Paragraph {
                paragraph_xml_bytes,
                ..
            } = self.phase
        {
            let next_paragraph_bytes = match paragraph_xml_bytes.checked_add(paragraph_bytes) {
                Some(value) => value,
                None => return Err(self.fail_overflow("paragraph XML bytes")),
            };
            self.phase = match self.phase {
                Phase::Paragraph {
                    run_open,
                    run_count,
                    run_text_bytes,
                    ..
                } => Phase::Paragraph {
                    run_open,
                    run_count,
                    run_text_bytes,
                    paragraph_xml_bytes: next_paragraph_bytes,
                },
                Phase::Ready => Phase::Ready,
            };
        }
        Ok(())
    }

    fn emit_escaped_text(&mut self, text: &str) -> Result<(), StreamingDocumentError> {
        let mut scratch = [0_u8; TEXT_SCRATCH_BYTES];
        let mut length = 0_usize;
        for character in text.chars() {
            let needed = escaped_character_len(character);
            let end = match length.checked_add(needed) {
                Some(value) => value,
                None => return Err(self.fail_overflow("escaped text bytes")),
            };
            if end > scratch.len() {
                let chunk_bytes = usize_to_u64(length, "escaped text bytes")
                    .map_err(|resource| self.fail_overflow(resource))?;
                self.context
                    .check()
                    .map_err(|error| self.fail_execution(error))?;
                self.emit_document(&scratch[..length], chunk_bytes, Some(chunk_bytes))?;
                length = 0;
            }
            let end = match length.checked_add(needed) {
                Some(value) => value,
                None => return Err(self.fail_overflow("escaped text bytes")),
            };
            let Some(destination) = scratch.get_mut(length..end) else {
                return Err(self.fail_invalid("escaped text exceeds fixed scratch"));
            };
            if append_character(character, destination) != needed {
                return Err(self.fail_invalid("escaped text did not fit fixed scratch"));
            }
            length = end;
        }
        if length != 0 {
            let chunk_bytes = usize_to_u64(length, "escaped text bytes")
                .map_err(|resource| self.fail_overflow(resource))?;
            self.context
                .check()
                .map_err(|error| self.fail_execution(error))?;
            self.emit_document(&scratch[..length], chunk_bytes, Some(chunk_bytes))?;
        }
        Ok(())
    }

    fn encoded_text_size(&mut self, text: &str) -> Result<(u64, u64), StreamingDocumentError> {
        let mut encoded = 0_u64;
        let mut characters = 0_u64;
        let mut pending_work = 0_u64;
        for character in text.chars() {
            if !is_plain_text_character(character) {
                return Err(self.fail_invalid(format!(
                    "character U+{:04X} requires structural WordprocessingML content",
                    u32::from(character)
                )));
            }
            let needed = usize_to_u64(escaped_character_len(character), "escaped text bytes")
                .map_err(|resource| self.fail_overflow(resource))?;
            encoded = checked_add(encoded, needed, "escaped text bytes")
                .map_err(|resource| self.fail_overflow(resource))?;
            characters = checked_add(characters, 1, "text characters")
                .map_err(|resource| self.fail_overflow(resource))?;
            pending_work = checked_add(pending_work, 1, "work")
                .map_err(|resource| self.fail_overflow(resource))?;
            if pending_work == TEXT_SCAN_CHECK_INTERVAL {
                self.context
                    .consume(Resource::Work, pending_work)
                    .map_err(|error| self.fail_execution(error))?;
                pending_work = 0;
            }
        }
        if pending_work != 0 {
            self.context
                .consume(Resource::Work, pending_work)
                .map_err(|error| self.fail_execution(error))?;
        }
        Ok((encoded, characters))
    }

    fn fail_io(&mut self, error: io::Error) -> StreamingDocumentError {
        let output = self.output_bytes();
        if let Some(execution) = execution_from_io_error(&error) {
            self.part.take();
            return self.fail_execution(execution);
        }
        if let Some(limit) = output_limit_from_io_error(&error) {
            self.part.take();
            return self.fail_limit(limit.resource, limit.observed, limit.limit);
        }
        if let Some(resource) = counter_overflow_from_io_error(&error) {
            self.part.take();
            return self.fail_overflow(resource);
        }
        if let Some(execution) = self.execution_state.get() {
            self.part.take();
            return self.fail_execution(execution);
        }
        if let Some(limit) = self.execution_state.output_limit() {
            self.part.take();
            return self.fail_limit("output bytes", limit.observed, limit.limit);
        }
        if let Some(resource) = self.execution_state.counter_overflow() {
            self.part.take();
            return self.fail_overflow(resource);
        }
        let kind = error.kind();
        let finish_error = self.part.take().and_then(|part| part.finish().err());
        let message = finish_error.as_ref().map_or_else(
            || error.to_string(),
            |finish_error| format!("{error}; package finalization: {finish_error}"),
        );
        let incomplete = output != 0
            || finish_error
                .as_ref()
                .is_some_and(|error| matches!(error, OpcError::IncompleteOutput { .. }));
        let source = Some(chained_error_source(error, finish_error));
        let fallback_message = message.clone();
        let reason = if incomplete {
            PoisonReason::IncompleteOutput { message, source }
        } else {
            PoisonReason::Io {
                kind,
                message,
                source,
            }
        };
        self.poison = Some(reason);
        self.poison.as_ref().map_or_else(
            || invalid_with_written(fallback_message, output),
            |reason| reason.error(output),
        )
    }
}

#[derive(Debug, Clone, Copy)]
struct RecordedOutputLimit {
    observed: u64,
    limit: u64,
}

#[derive(Default)]
struct ExecutionFailureState {
    error: Mutex<Option<ExecutionError>>,
    output_limit: Mutex<Option<RecordedOutputLimit>>,
    counter_overflow: Mutex<Option<&'static str>>,
}

impl ExecutionFailureState {
    fn record(&self, error: ExecutionError) {
        let mut recorded = self
            .error
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if recorded.is_none() {
            *recorded = Some(error);
        }
    }

    fn get(&self) -> Option<ExecutionError> {
        self.error
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .clone()
    }

    fn record_output_limit(&self, observed: u64, limit: u64) {
        let mut recorded = self
            .output_limit
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if recorded.is_none() {
            *recorded = Some(RecordedOutputLimit { observed, limit });
        }
    }

    fn output_limit(&self) -> Option<RecordedOutputLimit> {
        *self
            .output_limit
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }

    fn record_counter_overflow(&self, resource: &'static str) {
        let mut recorded = self
            .counter_overflow
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if recorded.is_none() {
            *recorded = Some(resource);
        }
    }

    fn counter_overflow(&self) -> Option<&'static str> {
        *self
            .counter_overflow
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }
}

struct BudgetedOutput<W> {
    writer: W,
    context: ExecutionContext,
    maximum: u64,
    accepted: u64,
    counter: Arc<AtomicU64>,
    execution_state: Arc<ExecutionFailureState>,
}

impl<W> BudgetedOutput<W> {
    fn new(
        writer: W,
        context: ExecutionContext,
        maximum: u64,
        counter: Arc<AtomicU64>,
        execution_state: Arc<ExecutionFailureState>,
    ) -> Self {
        Self {
            writer,
            context,
            maximum,
            accepted: 0,
            counter,
            execution_state,
        }
    }

    fn into_inner(self) -> io::Result<W> {
        Ok(self.writer)
    }
}

impl<W: Write> Write for BudgetedOutput<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.is_empty() {
            return Ok(0);
        }
        if let Err(error) = self.context.check() {
            self.execution_state.record(error.clone());
            return Err(execution_io_error(error));
        }
        let remaining = self.maximum.checked_sub(self.accepted).ok_or_else(|| {
            let error = counter_overflow_io_error("output bytes");
            self.execution_state.record_counter_overflow("output bytes");
            error
        })?;
        if remaining == 0 {
            let observed = self.accepted.checked_add(1).ok_or_else(|| {
                let error = counter_overflow_io_error("output bytes");
                self.execution_state.record_counter_overflow("output bytes");
                error
            })?;
            self.execution_state
                .record_output_limit(observed, self.maximum);
            return Err(output_limit_io_error(observed, self.maximum));
        }
        let requested_bytes = match u64::try_from(bytes.len()) {
            Ok(value) => value,
            Err(_error) => {
                self.execution_state
                    .record_counter_overflow("output request bytes");
                return Err(counter_overflow_io_error("output request bytes"));
            },
        };
        let requested = requested_bytes.min(remaining);
        let request_len = match usize::try_from(requested) {
            Ok(value) => value,
            Err(_error) => {
                self.execution_state.record_counter_overflow("output bytes");
                return Err(counter_overflow_io_error("output bytes"));
            },
        };
        let reservation = self
            .context
            .reserve(Resource::OutputBytes, requested)
            .map_err(|error| {
                self.execution_state.record(error.clone());
                execution_io_error(error)
            })?;
        let written = loop {
            if let Err(error) = self.context.check() {
                self.execution_state.record(error.clone());
                return Err(execution_io_error(error));
            }
            match self.writer.write(&bytes[..request_len]) {
                Err(error) if error.kind() == io::ErrorKind::Interrupted => continue,
                result => break result,
            }
        }?;
        let written_u64 = match u64::try_from(written) {
            Ok(value) => value,
            Err(_error) => {
                self.execution_state.record_counter_overflow("output bytes");
                return Err(counter_overflow_io_error("output bytes"));
            },
        };
        if written > request_len || !reservation.commit(written_u64) {
            return Err(io::Error::other(
                "streaming DOCX sink reported excess progress",
            ));
        }
        self.accepted = match self.accepted.checked_add(written_u64) {
            Some(value) => value,
            None => {
                self.execution_state.record_counter_overflow("output bytes");
                return Err(counter_overflow_io_error("output bytes"));
            },
        };
        self.counter.store(self.accepted, Ordering::Release);
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        loop {
            if let Err(error) = self.context.check() {
                self.execution_state.record(error.clone());
                return Err(execution_io_error(error));
            }
            match self.writer.flush() {
                Err(error) if error.kind() == io::ErrorKind::Interrupted => continue,
                Ok(()) => {
                    if let Err(error) = self.context.check() {
                        self.execution_state.record(error.clone());
                        return Err(execution_io_error(error));
                    }
                    return Ok(());
                },
                Err(error) => return Err(error),
            }
        }
    }
}

#[derive(Debug)]
struct ExecutionMarker(ExecutionError);

impl fmt::Display for ExecutionMarker {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl std::error::Error for ExecutionMarker {}

#[derive(Debug)]
struct OutputLimitMarker {
    observed: u64,
    limit: u64,
}

impl fmt::Display for OutputLimitMarker {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "streaming DOCX output limit exceeded: observed {}, limit {}",
            self.observed, self.limit
        )
    }
}

impl std::error::Error for OutputLimitMarker {}

#[derive(Debug)]
struct CounterOverflowMarker {
    resource: &'static str,
}

impl fmt::Display for CounterOverflowMarker {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "streaming DOCX counter overflowed for {}",
            self.resource
        )
    }
}

impl std::error::Error for CounterOverflowMarker {}

fn execution_io_error(error: ExecutionError) -> io::Error {
    io::Error::other(ExecutionMarker(error))
}

fn execution_from_io_error(error: &io::Error) -> Option<ExecutionError> {
    error
        .get_ref()
        .and_then(|source| source.downcast_ref::<ExecutionMarker>())
        .map(|marker| marker.0.clone())
}

struct OutputLimit {
    resource: &'static str,
    observed: u64,
    limit: u64,
}

fn output_limit_io_error(observed: u64, limit: u64) -> io::Error {
    io::Error::other(OutputLimitMarker { observed, limit })
}

fn output_limit_from_io_error(error: &io::Error) -> Option<OutputLimit> {
    error
        .get_ref()
        .and_then(|source| source.downcast_ref::<OutputLimitMarker>())
        .map(|marker| OutputLimit {
            resource: "output bytes",
            observed: marker.observed,
            limit: marker.limit,
        })
}

fn counter_overflow_io_error(resource: &'static str) -> io::Error {
    io::Error::other(CounterOverflowMarker { resource })
}

fn counter_overflow_from_io_error(error: &io::Error) -> Option<&'static str> {
    error
        .get_ref()
        .and_then(|source| source.downcast_ref::<CounterOverflowMarker>())
        .map(|marker| marker.resource)
}

fn validate_limits(limits: StreamingDocumentLimits) -> Result<(), StreamingDocumentError> {
    if limits.max_output_bytes < MIN_OUTPUT_BYTES {
        return Err(StreamingDocumentError::InvalidInput {
            message: "streaming DOCX output limit is too small for ZIP".to_owned(),
            written: 0,
        });
    }
    if limits.max_output_bytes > OPC_MAX_OUTPUT_BYTES {
        return Err(StreamingDocumentError::InvalidInput {
            message: format!(
                "streaming DOCX output limit exceeds the fixed OPC ZIP ceiling of {OPC_MAX_OUTPUT_BYTES} bytes"
            ),
            written: 0,
        });
    }
    if limits.max_document_xml_bytes > OPC_MAX_ENTRY_BYTES
        || limits.max_paragraph_xml_bytes > OPC_MAX_ENTRY_BYTES
    {
        return Err(StreamingDocumentError::InvalidInput {
            message: format!(
                "streaming DOCX XML limits exceed the fixed OPC entry ceiling of {OPC_MAX_ENTRY_BYTES} bytes"
            ),
            written: 0,
        });
    }
    if limits.max_scratch_bytes < FIXED_SCRATCH_BYTES {
        return Err(StreamingDocumentError::InvalidInput {
            message: format!(
                "streaming DOCX scratch limit must be at least {FIXED_SCRATCH_BYTES} bytes"
            ),
            written: 0,
        });
    }
    usize::try_from(limits.max_scratch_bytes).map_err(|_error| {
        StreamingDocumentError::InvalidInput {
            message: "streaming DOCX scratch limit exceeds usize".to_owned(),
            written: 0,
        }
    })?;
    let envelope =
        u64::try_from(DOCUMENT_PREFIX.len() + DOCUMENT_SUFFIX.len()).map_err(|_error| {
            StreamingDocumentError::InvalidInput {
                message: "streaming DOCX document XML envelope exceeds u64".to_owned(),
                written: 0,
            }
        })?;
    if limits.max_document_xml_bytes < envelope {
        return Err(StreamingDocumentError::InvalidInput {
            message: "streaming DOCX document XML limit is smaller than its envelope".to_owned(),
            written: 0,
        });
    }
    let minimum_paragraph = u64::try_from(PARAGRAPH_PREFIX.len() + PARAGRAPH_SUFFIX.len())
        .map_err(|_error| StreamingDocumentError::InvalidInput {
            message: "streaming DOCX paragraph envelope exceeds u64".to_owned(),
            written: 0,
        })?;
    if limits.max_paragraph_xml_bytes < minimum_paragraph {
        return Err(StreamingDocumentError::InvalidInput {
            message: "streaming DOCX paragraph XML limit is smaller than an empty paragraph"
                .to_owned(),
            written: 0,
        });
    }
    Ok(())
}

#[cfg(test)]
fn encoded_text_size(text: &str) -> Result<(u64, u64), String> {
    let mut encoded = 0_u64;
    let mut characters = 0_u64;
    for character in text.chars() {
        if !is_plain_text_character(character) {
            return Err(format!(
                "character U+{:04X} requires structural WordprocessingML content",
                u32::from(character)
            ));
        }
        encoded = encoded
            .checked_add(u64::try_from(escaped_character_len(character)).unwrap_or(u64::MAX))
            .ok_or_else(|| "escaped text byte count overflowed".to_owned())?;
        characters = characters
            .checked_add(1)
            .ok_or_else(|| "text character count overflowed".to_owned())?;
    }
    Ok((encoded, characters))
}

fn is_plain_text_character(character: char) -> bool {
    matches!(
        character,
        '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
    )
}

fn escaped_character_len(character: char) -> usize {
    match character {
        '&' => 5,
        '<' | '>' => 4,
        _ => character.len_utf8(),
    }
}

fn append_character(character: char, destination: &mut [u8]) -> usize {
    match character {
        '&' => copy_bytes(destination, b"&amp;"),
        '<' => copy_bytes(destination, b"&lt;"),
        '>' => copy_bytes(destination, b"&gt;"),
        _ => character.encode_utf8(destination).len(),
    }
}

fn copy_bytes(destination: &mut [u8], source: &[u8]) -> usize {
    if let Some(target) = destination.get_mut(..source.len()) {
        target.copy_from_slice(source);
        source.len()
    } else {
        0
    }
}

fn checked_increment(value: u64, resource: &'static str) -> Result<u64, &'static str> {
    value.checked_add(1).ok_or(resource)
}

fn checked_add(left: u64, right: u64, resource: &'static str) -> Result<u64, &'static str> {
    left.checked_add(right).ok_or(resource)
}

fn usize_to_u64(value: usize, resource: &'static str) -> Result<u64, &'static str> {
    u64::try_from(value).map_err(|_error| resource)
}

fn resource_name(resource: Resource) -> &'static str {
    match resource {
        Resource::Memory => "memory",
        Resource::InputBytes => "input bytes",
        Resource::OutputBytes => "output bytes",
        Resource::Objects => "objects",
        Resource::Depth => "depth",
        Resource::Work => "work",
        _ => "unknown",
    }
}

fn poison_from_execution(error: ExecutionError) -> PoisonReason {
    match error {
        ExecutionError::Cancelled => PoisonReason::Cancelled,
        ExecutionError::ResourceLimit(limit) => PoisonReason::Limit {
            resource: resource_name(limit.resource),
            observed: limit.observed,
            limit: limit.limit,
        },
        other => PoisonReason::InvalidInput(other.to_string()),
    }
}

fn map_execution(error: ExecutionError, written: u64) -> StreamingDocumentError {
    poison_from_execution(error).error(written)
}

fn invalid_with_written(message: String, written: u64) -> StreamingDocumentError {
    StreamingDocumentError::InvalidInput { message, written }
}

fn map_opc(error: OpcError, written: u64) -> StreamingDocumentError {
    let incomplete = matches!(error, OpcError::IncompleteOutput { .. });
    let source = error_source(error);
    let message = source.to_string();
    if incomplete {
        StreamingDocumentError::IncompleteOutput {
            message,
            written,
            source: Some(source),
        }
    } else {
        StreamingDocumentError::Opc {
            message,
            written,
            source: Some(source),
        }
    }
}

fn map_opc_with_state(
    error: OpcError,
    written: u64,
    execution_state: &ExecutionFailureState,
) -> StreamingDocumentError {
    if let Some(error) = execution_state.get() {
        return map_execution(error, written);
    }
    if let Some(limit) = execution_state.output_limit() {
        return StreamingDocumentError::LimitExceeded {
            resource: "output bytes",
            observed: limit.observed,
            limit: limit.limit,
            written,
        };
    }
    map_opc(error, written)
}

fn map_member_write_error(
    write_error: io::Error,
    finish_error: Option<OpcError>,
    output: u64,
    execution_state: &ExecutionFailureState,
) -> StreamingDocumentError {
    if let Some(execution) = execution_from_io_error(&write_error) {
        return map_execution(execution, output);
    }
    if let Some(limit) = output_limit_from_io_error(&write_error) {
        return StreamingDocumentError::LimitExceeded {
            resource: limit.resource,
            observed: limit.observed,
            limit: limit.limit,
            written: output,
        };
    }
    if let Some(resource) = counter_overflow_from_io_error(&write_error) {
        return StreamingDocumentError::CounterOverflow {
            resource,
            written: output,
        };
    }
    if let Some(execution) = execution_state.get() {
        return map_execution(execution, output);
    }
    if let Some(limit) = execution_state.output_limit() {
        return StreamingDocumentError::LimitExceeded {
            resource: "output bytes",
            observed: limit.observed,
            limit: limit.limit,
            written: output,
        };
    }
    if let Some(resource) = execution_state.counter_overflow() {
        return StreamingDocumentError::CounterOverflow {
            resource,
            written: output,
        };
    }
    let kind = write_error.kind();
    let message = finish_error.as_ref().map_or_else(
        || write_error.to_string(),
        |finish_error| format!("{write_error}; package finalization: {finish_error}"),
    );
    if let Some(finish_error) = finish_error {
        let incomplete = matches!(finish_error, OpcError::IncompleteOutput { .. });
        let source = chained_error_source(write_error, Some(finish_error));
        if incomplete || output != 0 {
            return StreamingDocumentError::IncompleteOutput {
                message,
                written: output,
                source: Some(source),
            };
        }
        return StreamingDocumentError::Opc {
            message,
            written: output,
            source: Some(source),
        };
    }
    let source = error_source(write_error);
    if output != 0 {
        StreamingDocumentError::IncompleteOutput {
            message,
            written: output,
            source: Some(source),
        }
    } else {
        StreamingDocumentError::Io {
            kind,
            message,
            written: output,
            source: Some(source),
        }
    }
}

fn poison_from_error(error: &StreamingDocumentError) -> PoisonReason {
    match error {
        StreamingDocumentError::Io {
            kind,
            message,
            source,
            ..
        } => PoisonReason::Io {
            kind: *kind,
            message: message.clone(),
            source: source.clone(),
        },
        StreamingDocumentError::InvalidInput { message, .. } => {
            PoisonReason::InvalidInput(message.clone())
        },
        StreamingDocumentError::LimitExceeded {
            resource,
            observed,
            limit,
            ..
        } => PoisonReason::Limit {
            resource,
            observed: *observed,
            limit: *limit,
        },
        StreamingDocumentError::Cancelled { .. } => PoisonReason::Cancelled,
        StreamingDocumentError::Opc {
            message, source, ..
        } => PoisonReason::Opc {
            message: message.clone(),
            source: source.clone(),
        },
        StreamingDocumentError::IncompleteOutput {
            message, source, ..
        } => PoisonReason::IncompleteOutput {
            message: message.clone(),
            source: source.clone(),
        },
        StreamingDocumentError::CounterOverflow { resource, .. } => {
            PoisonReason::CounterOverflow(resource)
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits};
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::atomic::AtomicBool;

    fn context_for(budget: Budget) -> (CancellationSource, ExecutionContext) {
        let (source, token) = CancellationSource::pair();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("one worker"),
            NonZeroUsize::new(1).expect("one task"),
            NonZeroU64::new(1024 * 1024).expect("nonzero in-flight bytes"),
            1,
        )
        .expect("valid execution limits");
        (source, ExecutionContext::new(budget, token, execution))
    }

    fn context_with_budget_limits(
        limits: Limits,
    ) -> (Budget, CancellationSource, ExecutionContext) {
        let budget = Budget::root("docx-stream-custom", limits);
        let (source, context) = context_for(budget.clone());
        (budget, source, context)
    }

    fn test_context() -> (Budget, CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "docx-stream-test",
            Limits::new(
                1024 * 1024,
                1024 * 1024,
                1024 * 1024,
                10_000,
                64,
                10_000_000,
            ),
        );
        let (source, context) = context_for(budget.clone());
        (budget, source, context)
    }

    fn limits() -> StreamingDocumentLimits {
        StreamingDocumentLimits::default()
    }

    fn write_sample<W: Write>(writer: &mut StreamingDocumentWriter<W>) {
        writer.start_paragraph().expect("paragraph");
        writer.start_run().expect("first run");
        writer.write_text(" leading & <é> ").expect("first text");
        writer.finish_run().expect("first run finish");
        writer.start_run().expect("second run");
        writer.write_text("tail").expect("second text");
        writer.finish_run().expect("second run finish");
        writer.finish_paragraph().expect("paragraph finish");
        writer.start_paragraph().expect("empty paragraph");
        writer.finish_paragraph().expect("empty paragraph finish");
    }

    fn render_sample() -> Vec<u8> {
        let (_budget, _source, context) = test_context();
        let mut writer =
            StreamingDocumentWriter::new(Vec::new(), context, limits()).expect("streaming writer");
        write_sample(&mut writer);
        writer.finish().expect("package finish")
    }

    #[test]
    fn plain_text_validation_and_escaping_are_preflighted() {
        assert_eq!(encoded_text_size("<&>é").expect("valid text"), (15, 4));
        for invalid in ["\t", "\n", "\r", "\0", "\u{000B}", "\u{FFFE}"] {
            let (_budget, _source, context) = test_context();
            let mut writer = StreamingDocumentWriter::new(Vec::new(), context, limits())
                .expect("streaming writer");
            writer.start_paragraph().expect("paragraph");
            writer.start_run().expect("run");
            let before_output = writer.output_bytes();
            let error = writer.write_text(invalid).expect_err("invalid plain text");
            assert!(matches!(error, StreamingDocumentError::InvalidInput { .. }));
            assert_eq!(error.written(), before_output);
            assert_eq!(writer.input_bytes(), 0);
            assert!(writer.is_poisoned());
            let repeated = writer
                .write_text("after poison")
                .expect_err("poisoned writer");
            assert_eq!(repeated.written(), before_output);
        }
    }

    #[test]
    fn state_errors_are_retryable_but_semantic_failures_poison() {
        let (_budget, _source, context) = test_context();
        let mut writer =
            StreamingDocumentWriter::new(Vec::new(), context, limits()).expect("streaming writer");
        assert!(matches!(
            writer.finish_run(),
            Err(StreamingDocumentError::InvalidInput { .. })
        ));
        assert!(!writer.is_poisoned());
        assert!(matches!(
            writer.start_run(),
            Err(StreamingDocumentError::InvalidInput { .. })
        ));
        assert!(!writer.is_poisoned());
        writer.start_paragraph().expect("paragraph");
        writer.start_run().expect("run");
        assert!(matches!(
            writer.start_run(),
            Err(StreamingDocumentError::InvalidInput { .. })
        ));
        assert!(!writer.is_poisoned());
        assert!(matches!(
            writer.finish_paragraph(),
            Err(StreamingDocumentError::InvalidInput { .. })
        ));
        assert!(!writer.is_poisoned());
        writer.finish_run().expect("run finish");
        writer.finish_paragraph().expect("paragraph finish");
        writer.finish().expect("package finish");
    }

    #[test]
    fn deterministic_package_is_bounded_to_three_members() {
        let first = render_sample();
        let second = render_sample();
        assert_eq!(first, second);
        let package = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(first.clone())
            .expect("physical package");
        assert_eq!(
            package.member_names().expect("member names"),
            vec![
                "[Content_Types].xml".to_owned(),
                "_rels/.rels".to_owned(),
                "word/document.xml".to_owned(),
            ]
        );
        let document_uri = PackURI::new(DOCUMENT_URI).expect("document URI");
        let document_xml = package.blob_for(&document_uri).expect("document XML");
        assert_eq!(
            document_xml,
            concat!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>",
                "<w:p><w:r><w:t xml:space=\"preserve\"> leading &amp; &lt;é&gt; </w:t></w:r>",
                "<w:r><w:t xml:space=\"preserve\">tail</w:t></w:r></w:p><w:p></w:p>",
                "<w:sectPr/></w:body></w:document>"
            )
            .as_bytes()
        );
        assert!(!document_xml.windows(2).any(|pair| pair == b"\r\n"));
    }

    #[test]
    fn escaped_chunk_boundaries_are_counted_once_at_exact_xml_limits() {
        let text = format!("{}&é<>{}", "a".repeat(58), "b".repeat(5));
        let (encoded, _) = encoded_text_size(&text).expect("boundary text");
        assert!(encoded > u64::try_from(TEXT_SCRATCH_BYTES).expect("scratch size"));
        let paragraph_bytes = u64::try_from(
            PARAGRAPH_PREFIX.len()
                + RUN_PREFIX.len()
                + usize::try_from(encoded).expect("encoded text fits usize")
                + RUN_SUFFIX.len()
                + PARAGRAPH_SUFFIX.len(),
        )
        .expect("paragraph bytes");
        let document_bytes = u64::try_from(
            DOCUMENT_PREFIX.len()
                + usize::try_from(paragraph_bytes).expect("paragraph fits usize")
                + DOCUMENT_SUFFIX.len(),
        )
        .expect("document bytes");
        let exact_limits = StreamingDocumentLimits {
            max_paragraph_xml_bytes: paragraph_bytes,
            max_document_xml_bytes: document_bytes,
            ..limits()
        };
        let (_budget, _source, context) = test_context();
        let mut writer = StreamingDocumentWriter::new(Vec::new(), context, exact_limits)
            .expect("exact boundary writer");
        writer.start_paragraph().expect("paragraph");
        writer.start_run().expect("run");
        writer.write_text(&text).expect("boundary text");
        let expected_before_close = u64::try_from(
            DOCUMENT_PREFIX.len()
                + PARAGRAPH_PREFIX.len()
                + RUN_PREFIX.len()
                + usize::try_from(encoded).expect("encoded fits usize"),
        )
        .expect("document progress");
        assert_eq!(writer.document_xml_bytes(), expected_before_close);
        writer.finish_run().expect("run finish");
        writer.finish_paragraph().expect("paragraph finish");
        writer.finish().expect("exact XML limits");

        let (_budget, _source, context) = test_context();
        let mut one_under_paragraph = StreamingDocumentWriter::new(
            Vec::new(),
            context,
            StreamingDocumentLimits {
                max_paragraph_xml_bytes: paragraph_bytes - 1,
                ..exact_limits
            },
        )
        .expect("one-under paragraph writer");
        one_under_paragraph.start_paragraph().expect("paragraph");
        one_under_paragraph.start_run().expect("run");
        one_under_paragraph
            .write_text(&text)
            .expect("boundary text");
        one_under_paragraph.finish_run().expect("run finish");
        let error = one_under_paragraph
            .finish_paragraph()
            .expect_err("one-under paragraph XML");
        assert!(matches!(
            error,
            StreamingDocumentError::LimitExceeded {
                resource: "paragraph XML bytes",
                observed,
                limit,
                ..
            } if observed == paragraph_bytes && limit == paragraph_bytes - 1
        ));

        let (_budget, _source, context) = test_context();
        let mut one_under_document = StreamingDocumentWriter::new(
            Vec::new(),
            context,
            StreamingDocumentLimits {
                max_document_xml_bytes: document_bytes - 1,
                ..exact_limits
            },
        )
        .expect("one-under document writer");
        one_under_document.start_paragraph().expect("paragraph");
        one_under_document.start_run().expect("run");
        one_under_document
            .write_text(&text)
            .expect("text before suffix");
        one_under_document.finish_run().expect("run finish");
        one_under_document
            .finish_paragraph()
            .expect("paragraph finish");
        let error = one_under_document
            .finish()
            .expect_err("one-under document XML");
        assert!(matches!(
            error,
            StreamingDocumentError::LimitExceeded {
                resource: "document XML bytes",
                observed,
                limit,
                ..
            } if observed == document_bytes && limit == document_bytes - 1
        ));
    }

    #[test]
    fn finite_limits_are_exact_and_one_over() {
        let (_budget, _source, context) = test_context();
        let mut paragraph_limited = StreamingDocumentWriter::new(
            Vec::new(),
            context,
            StreamingDocumentLimits {
                max_paragraphs: 0,
                ..limits()
            },
        )
        .expect("writer");
        let error = paragraph_limited
            .start_paragraph()
            .expect_err("paragraph limit");
        assert!(matches!(
            error,
            StreamingDocumentError::LimitExceeded {
                resource: "paragraphs",
                observed: 1,
                limit: 0,
                ..
            }
        ));
        assert!(paragraph_limited.is_poisoned());

        let (_budget, _source, context) = test_context();
        let mut run_limited = StreamingDocumentWriter::new(
            Vec::new(),
            context,
            StreamingDocumentLimits {
                max_runs: 1,
                ..limits()
            },
        )
        .expect("writer");
        run_limited.start_paragraph().expect("paragraph");
        run_limited.start_run().expect("first run");
        run_limited.finish_run().expect("first run finish");
        let error = run_limited.start_run().expect_err("run limit");
        assert!(matches!(
            error,
            StreamingDocumentError::LimitExceeded {
                resource: "runs",
                observed: 2,
                limit: 1,
                ..
            }
        ));

        let (_budget, _source, context) = test_context();
        let mut input_limited = StreamingDocumentWriter::new(
            Vec::new(),
            context,
            StreamingDocumentLimits {
                max_input_bytes: 3,
                max_run_text_bytes: 3,
                ..limits()
            },
        )
        .expect("writer");
        input_limited.start_paragraph().expect("paragraph");
        input_limited.start_run().expect("run");
        input_limited.write_text("abc").expect("exact input");
        let error = input_limited.write_text("d").expect_err("one over input");
        assert!(matches!(
            error,
            StreamingDocumentError::LimitExceeded {
                resource: "run text bytes",
                ..
            }
        ));
        assert_eq!(input_limited.input_bytes(), 3);
        assert_eq!(input_limited.output_bytes(), error.written());
    }

    #[test]
    fn output_limit_is_exact_and_one_over() {
        let baseline = render_sample();
        let exact_output = u64::try_from(baseline.len()).expect("output length");

        let (_budget, _source, context) = test_context();
        let mut exact = StreamingDocumentWriter::new(
            Vec::new(),
            context,
            StreamingDocumentLimits {
                max_output_bytes: exact_output,
                ..limits()
            },
        )
        .expect("exact output writer");
        write_sample(&mut exact);
        let exact_bytes = exact.finish().expect("exact output");
        assert_eq!(
            u64::try_from(exact_bytes.len()).expect("exact bytes"),
            exact_output
        );

        let (_budget, _source, context) = test_context();
        let mut over = StreamingDocumentWriter::new(
            Vec::new(),
            context,
            StreamingDocumentLimits {
                max_output_bytes: exact_output.saturating_sub(1),
                ..limits()
            },
        )
        .expect("one-under output writer");
        write_sample(&mut over);
        let error = match over.finish() {
            Ok(_) => panic!("one-under output unexpectedly finished"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            StreamingDocumentError::LimitExceeded {
                resource: "output bytes",
                observed,
                limit,
                written,
            } if observed == exact_output && limit == exact_output - 1 && written == exact_output - 1
        ));

        let (_budget, _source, context) = test_context();
        let mut above = StreamingDocumentWriter::new(
            Vec::new(),
            context,
            StreamingDocumentLimits {
                max_output_bytes: exact_output.saturating_add(1),
                ..limits()
            },
        )
        .expect("one-above output writer");
        write_sample(&mut above);
        above.finish().expect("one-above output");
    }

    #[test]
    fn hierarchical_input_and_output_budgets_have_exact_boundaries() {
        let input_budget = Budget::root(
            "docx-stream-input-parent",
            Limits::new(1024 * 1024, 3, 1024 * 1024, 100, 64, 1_000_000),
        )
        .child(
            "docx-stream-input-child",
            Limits::new(1024 * 1024, 3, 1024 * 1024, 100, 64, 1_000_000),
        );
        let (_source, context) = context_for(input_budget.clone());
        let mut exact_input = StreamingDocumentWriter::new(Vec::new(), context, limits())
            .expect("input budget writer");
        exact_input.start_paragraph().expect("paragraph");
        exact_input.start_run().expect("run");
        exact_input.write_text("abc").expect("exact input budget");
        assert_eq!(input_budget.used(Resource::InputBytes), 3);
        let error = exact_input
            .write_text("d")
            .expect_err("one-over input budget");
        assert!(matches!(
            error,
            StreamingDocumentError::LimitExceeded {
                resource: "input bytes",
                observed: 4,
                limit: 3,
                ..
            }
        ));

        let baseline = render_sample();
        let output_limit = u64::try_from(baseline.len()).expect("baseline output");
        let output_budget = Budget::root(
            "docx-stream-output-parent",
            Limits::new(
                1024 * 1024,
                1024 * 1024,
                output_limit,
                10_000,
                64,
                10_000_000,
            ),
        )
        .child(
            "docx-stream-output-child",
            Limits::new(
                1024 * 1024,
                1024 * 1024,
                output_limit,
                10_000,
                64,
                10_000_000,
            ),
        );
        let (_source, context) = context_for(output_budget.clone());
        let mut exact_output = StreamingDocumentWriter::new(
            Vec::new(),
            context,
            StreamingDocumentLimits {
                max_output_bytes: output_limit,
                ..limits()
            },
        )
        .expect("output budget writer");
        write_sample(&mut exact_output);
        let bytes = exact_output.finish().expect("exact output budget");
        assert_eq!(
            u64::try_from(bytes.len()).expect("output bytes"),
            output_limit
        );
        assert_eq!(output_budget.used(Resource::OutputBytes), output_limit);

        let output_under = Budget::root(
            "docx-stream-output-under-parent",
            Limits::new(
                1024 * 1024,
                1024 * 1024,
                output_limit - 1,
                10_000,
                64,
                10_000_000,
            ),
        )
        .child(
            "docx-stream-output-under-child",
            Limits::new(
                1024 * 1024,
                1024 * 1024,
                output_limit - 1,
                10_000,
                64,
                10_000_000,
            ),
        );
        let (_source, context) = context_for(output_under);
        let mut output_under_writer = StreamingDocumentWriter::new(
            Vec::new(),
            context,
            StreamingDocumentLimits {
                max_output_bytes: output_limit,
                ..limits()
            },
        )
        .expect("one-under output budget initial publication");
        write_sample(&mut output_under_writer);
        let error = output_under_writer
            .finish()
            .expect_err("one-under output budget");
        assert!(matches!(
            error,
            StreamingDocumentError::LimitExceeded {
                resource: "output bytes",
                ..
            }
        ));
    }

    #[test]
    fn counter_overflow_is_reported_as_a_finite_limit_failure() {
        let (_budget, _source, context) = test_context();
        let mut input_overflow = StreamingDocumentWriter::new(Vec::new(), context, limits())
            .expect("input overflow writer");
        input_overflow.start_paragraph().expect("paragraph");
        input_overflow.start_run().expect("run");
        input_overflow.input_bytes = u64::MAX;
        let error = input_overflow
            .write_text("x")
            .expect_err("input counter overflow");
        assert!(matches!(
            error,
            StreamingDocumentError::CounterOverflow {
                resource: "input bytes",
                ..
            }
        ));

        let (_budget, _source, context) = test_context();
        let mut document_overflow = StreamingDocumentWriter::new(Vec::new(), context, limits())
            .expect("document overflow writer");
        document_overflow.document_xml_bytes = u64::MAX;
        let error = match document_overflow.finish() {
            Ok(_) => panic!("document counter overflow unexpectedly finished"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            StreamingDocumentError::CounterOverflow {
                resource: "document XML bytes",
                ..
            }
        ));
    }

    #[test]
    fn fixed_memory_reservation_releases_through_hierarchy() {
        let parent = Budget::root(
            "docx-stream-parent",
            Limits::new(
                FIXED_SCRATCH_BYTES,
                1024 * 1024,
                1024 * 1024,
                10_000,
                64,
                10_000,
            ),
        );
        let child = parent.child(
            "docx-stream-child",
            Limits::new(
                FIXED_SCRATCH_BYTES,
                1024 * 1024,
                1024 * 1024,
                10_000,
                64,
                10_000,
            ),
        );
        let (_source, context) = context_for(child.clone());
        let writer = StreamingDocumentWriter::new(Vec::new(), context, limits())
            .expect("exact scratch reservation");
        assert_eq!(child.used(Resource::Memory), FIXED_SCRATCH_BYTES);
        assert_eq!(parent.used(Resource::Memory), FIXED_SCRATCH_BYTES);
        drop(writer);
        assert_eq!(child.used(Resource::Memory), 0);
        assert_eq!(parent.used(Resource::Memory), 0);

        let under = Budget::root(
            "docx-stream-under",
            Limits::new(
                FIXED_SCRATCH_BYTES - 1,
                1024 * 1024,
                1024 * 1024,
                10_000,
                64,
                10_000,
            ),
        );
        let (_source, context) = context_for(under);
        let error = match StreamingDocumentWriter::new(Vec::new(), context, limits()) {
            Ok(_) => panic!("one-under scratch reservation unexpectedly succeeded"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            StreamingDocumentError::LimitExceeded {
                resource: "memory",
                ..
            }
        ));
    }

    #[test]
    fn fixed_objects_and_work_are_preflighted_before_first_output() {
        let exact_limits = Limits::new(
            1024 * 1024,
            1024 * 1024,
            1024 * 1024,
            FIXED_OBJECT_CHARGE,
            64,
            FIXED_WORK_CHARGE,
        );
        let (budget, _source, context) = context_with_budget_limits(exact_limits);
        let writer = StreamingDocumentWriter::new(Vec::new(), context, limits())
            .expect("exact fixed charges");
        assert!(writer.output_bytes() > 0);
        assert_eq!(budget.used(Resource::Objects), FIXED_OBJECT_CHARGE);
        assert_eq!(budget.used(Resource::Work), FIXED_WORK_CHARGE);
        drop(writer);

        for (resource, label) in [(Resource::Objects, "objects"), (Resource::Work, "work")] {
            let object_limit = if resource == Resource::Objects {
                FIXED_OBJECT_CHARGE - 1
            } else {
                FIXED_OBJECT_CHARGE + 1
            };
            let work_limit = if resource == Resource::Work {
                FIXED_WORK_CHARGE - 1
            } else {
                FIXED_WORK_CHARGE + 1
            };
            let (_budget, _source, context) = context_with_budget_limits(Limits::new(
                1024 * 1024,
                1024 * 1024,
                1024 * 1024,
                object_limit,
                64,
                work_limit,
            ));
            let error = match StreamingDocumentWriter::new(Vec::new(), context, limits()) {
                Ok(_) => panic!("one-under fixed charge unexpectedly succeeded"),
                Err(error) => error,
            };
            assert!(matches!(
                error,
                StreamingDocumentError::LimitExceeded { resource: actual, written: 0, .. }
                    if actual == label
            ));
        }
    }

    #[test]
    fn hidden_opc_ceilings_are_rejected_instead_of_being_opaque_runtime_limits() {
        let (_budget, _source, context) = test_context();
        let error = match StreamingDocumentWriter::new(
            Vec::new(),
            context,
            StreamingDocumentLimits {
                max_output_bytes: OPC_MAX_OUTPUT_BYTES + 1,
                ..limits()
            },
        ) {
            Ok(_) => panic!("output ceiling unexpectedly accepted"),
            Err(error) => error,
        };
        assert!(matches!(error, StreamingDocumentError::InvalidInput { .. }));

        let (_budget, _source, context) = test_context();
        let error = match StreamingDocumentWriter::new(
            Vec::new(),
            context,
            StreamingDocumentLimits {
                max_paragraph_xml_bytes: OPC_MAX_ENTRY_BYTES + 1,
                ..limits()
            },
        ) {
            Ok(_) => panic!("entry ceiling unexpectedly accepted"),
            Err(error) => error,
        };
        assert!(matches!(error, StreamingDocumentError::InvalidInput { .. }));
    }

    #[test]
    fn text_size_scan_charges_work_at_a_bounded_checkpoint() {
        let work_limit = FIXED_WORK_CHARGE + 2 + TEXT_SCAN_CHECK_INTERVAL - 1;
        let (_budget, _source, context) = context_with_budget_limits(Limits::new(
            1024 * 1024,
            1024 * 1024,
            1024 * 1024,
            10_000,
            64,
            work_limit,
        ));
        let mut writer =
            StreamingDocumentWriter::new(Vec::new(), context, limits()).expect("checkpoint writer");
        writer.start_paragraph().expect("paragraph");
        writer.start_run().expect("run");
        let before = writer.document_xml_bytes();
        let error = writer
            .write_text(&"a".repeat(TEXT_SCAN_CHECK_INTERVAL as usize))
            .expect_err("work checkpoint limit");
        assert!(matches!(
            error,
            StreamingDocumentError::LimitExceeded {
                resource: "work",
                observed,
                limit,
                ..
            } if observed == work_limit + 1 && limit == work_limit
        ));
        assert_eq!(writer.document_xml_bytes(), before);
    }

    #[derive(Debug, Default)]
    struct ShortSink {
        bytes: Vec<u8>,
        max_chunk: usize,
    }

    impl Write for ShortSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            let count = bytes.len().min(self.max_chunk.max(1));
            self.bytes.extend_from_slice(&bytes[..count]);
            Ok(count)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[derive(Debug, Default)]
    struct InterruptedSink {
        bytes: Vec<u8>,
        interrupted: bool,
        flush_interrupted: bool,
    }

    impl Write for InterruptedSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if !self.interrupted {
                self.interrupted = true;
                return Err(io::Error::new(io::ErrorKind::Interrupted, "retry"));
            }
            self.bytes.extend_from_slice(bytes);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            if !self.flush_interrupted {
                self.flush_interrupted = true;
                return Err(io::Error::new(io::ErrorKind::Interrupted, "retry flush"));
            }
            Ok(())
        }
    }

    #[derive(Debug, Default)]
    struct ZeroSink;

    impl Write for ZeroSink {
        fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
            Ok(0)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[derive(Debug)]
    struct FailingSink {
        bytes: Vec<u8>,
        fail_after: usize,
    }

    impl Write for FailingSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if self.bytes.len() >= self.fail_after {
                return Err(io::Error::new(io::ErrorKind::BrokenPipe, "failed"));
            }
            let count = bytes.len().min(self.fail_after - self.bytes.len());
            self.bytes.extend_from_slice(&bytes[..count]);
            Ok(count)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn non_seek_short_and_interrupted_sinks_finish_without_buffering() {
        let (_budget, _source, context) = test_context();
        let mut writer = StreamingDocumentWriter::new(
            ShortSink {
                max_chunk: 2,
                ..ShortSink::default()
            },
            context,
            limits(),
        )
        .expect("short sink writer");
        write_sample(&mut writer);
        let short = writer.finish().expect("short sink finish");
        assert!(!short.bytes.is_empty());

        let (_budget, _source, context) = test_context();
        let mut writer =
            StreamingDocumentWriter::new(InterruptedSink::default(), context, limits())
                .expect("interrupted sink writer");
        write_sample(&mut writer);
        let interrupted = writer.finish().expect("interrupted sink finish");
        assert!(!interrupted.bytes.is_empty());
        assert!(interrupted.flush_interrupted);
    }

    #[test]
    fn zero_and_failing_sinks_report_typed_progress_and_poisoning() {
        let (_budget, _source, context) = test_context();
        let error = match StreamingDocumentWriter::new(ZeroSink, context, limits()) {
            Ok(_) => panic!("zero sink unexpectedly succeeded"),
            Err(error) => error,
        };
        assert!(
            matches!(&error, StreamingDocumentError::Opc { written: 0, .. }),
            "unexpected zero-sink error: {error:?}"
        );

        let (_budget, _source, baseline_context) = test_context();
        let baseline = StreamingDocumentWriter::new(Vec::new(), baseline_context, limits())
            .expect("baseline writer");
        let prefix = baseline.output_bytes();
        drop(baseline);

        let (_budget, _source, context) = test_context();
        let mut writer = StreamingDocumentWriter::new(
            FailingSink {
                bytes: Vec::new(),
                fail_after: usize::try_from(prefix.saturating_add(1)).expect("prefix fits usize"),
            },
            context,
            limits(),
        )
        .expect("failing sink writer");
        assert_eq!(writer.output_bytes(), prefix);
        writer.start_paragraph().expect("paragraph");
        writer.finish_paragraph().expect("paragraph finish");
        let error = match writer.finish() {
            Ok(_) => panic!("failing sink unexpectedly finished"),
            Err(error) => error,
        };
        assert!(
            matches!(
                &error,
                StreamingDocumentError::IncompleteOutput {
                    written,
                    ..
                } if *written == prefix + 1
            ),
            "unexpected failing-sink error: {error:?}, prefix={prefix}"
        );
        assert_eq!(error.written(), prefix + 1);
        assert!(std::error::Error::source(&error).is_some());
    }

    #[test]
    fn controlled_sink_preserves_document_progress_through_deflate_failure() {
        let (_budget, _source, baseline_context) = test_context();
        let baseline = StreamingDocumentWriter::new(Vec::new(), baseline_context, limits())
            .expect("baseline writer");
        let prefix = baseline.output_bytes();
        drop(baseline);

        let (_budget, _source, context) = test_context();
        let mut writer = StreamingDocumentWriter::new(
            FailingSink {
                bytes: Vec::new(),
                fail_after: usize::try_from(prefix + 1).expect("prefix fits usize"),
            },
            context,
            limits(),
        )
        .expect("controlled failing writer");
        let chunk = vec![b'a'; 1024 * 1024];
        let chunk_bytes = u64::try_from(chunk.len()).expect("chunk bytes");
        let mut attempts = 0_u32;
        let error = loop {
            attempts += 1;
            let before = writer.document_xml_bytes();
            let result = writer.emit_document(&chunk, chunk_bytes, None);
            let after = writer.document_xml_bytes();
            if result.is_ok() {
                assert_eq!(after, before + chunk_bytes);
                assert!(attempts < 128, "controlled sink did not fail");
                continue;
            }
            let error = result.expect_err("controlled sink failure");
            assert!(after >= before);
            assert!(after - before <= chunk_bytes);
            break error;
        };
        assert!(matches!(
            error,
            StreamingDocumentError::IncompleteOutput { .. }
        ));
        assert_eq!(error.written(), prefix + 1);
        let source = std::error::Error::source(&error).expect("sink source");
        assert!(source.to_string().contains("failed"));
        assert!(source.source().is_some(), "OPC finalization source");
    }

    #[test]
    fn cancellation_is_observed_before_next_output() {
        let (_budget, source, context) = test_context();
        let mut writer =
            StreamingDocumentWriter::new(Vec::new(), context, limits()).expect("streaming writer");
        let prefix = writer.output_bytes();
        source.cancel();
        let error = writer.start_paragraph().expect_err("cancelled writer");
        assert!(matches!(
            error,
            StreamingDocumentError::Cancelled { written } if written == prefix
        ));
        assert_eq!(writer.output_bytes(), prefix);
        assert!(writer.is_poisoned());
    }

    #[derive(Debug)]
    struct CancellingFlushSink {
        bytes: Vec<u8>,
        source: CancellationSource,
        armed: Arc<AtomicBool>,
        interrupted: bool,
    }

    impl Write for CancellingFlushSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            self.bytes.extend_from_slice(bytes);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            if self.armed.load(Ordering::Acquire) && !self.interrupted {
                self.interrupted = true;
                self.source.cancel();
                return Err(io::Error::new(
                    io::ErrorKind::Interrupted,
                    "cancel during flush retry",
                ));
            }
            Ok(())
        }
    }

    #[test]
    fn cancellation_during_flush_retry_is_checked_before_retrying() {
        let (_budget, source, context) = test_context();
        let armed = Arc::new(AtomicBool::new(false));
        let mut writer = StreamingDocumentWriter::new(
            CancellingFlushSink {
                bytes: Vec::new(),
                source,
                armed: Arc::clone(&armed),
                interrupted: false,
            },
            context,
            limits(),
        )
        .expect("cancelling flush writer");
        armed.store(true, Ordering::Release);
        writer.start_paragraph().expect("paragraph");
        writer.finish_paragraph().expect("paragraph finish");
        let error = writer
            .finish()
            .expect_err("flush cancellation should poison publication");
        assert!(matches!(error, StreamingDocumentError::Cancelled { .. }));
    }

    #[derive(Debug)]
    struct CancellingSink {
        bytes: Vec<u8>,
        source: CancellationSource,
        armed: Arc<AtomicBool>,
        cancelled: bool,
        accepted: Arc<AtomicU64>,
    }

    impl Write for CancellingSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            self.bytes.extend_from_slice(bytes);
            self.accepted.fetch_add(
                u64::try_from(bytes.len()).unwrap_or(u64::MAX),
                Ordering::AcqRel,
            );
            if self.armed.load(Ordering::Acquire) && !self.cancelled {
                self.cancelled = true;
                self.source.cancel();
            }
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn cancellation_during_package_finalization_preserves_typed_progress() {
        let budget = Budget::root(
            "docx-stream-finalization-cancel",
            Limits::new(
                1024 * 1024,
                1024 * 1024,
                1024 * 1024,
                10_000,
                64,
                10_000_000,
            ),
        );
        let (source, token) = CancellationSource::pair();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("one worker"),
            NonZeroUsize::new(1).expect("one task"),
            NonZeroU64::new(1024 * 1024).expect("in-flight bytes"),
            1,
        )
        .expect("valid execution limits");
        let context = ExecutionContext::new(budget, token, execution);
        let armed = Arc::new(AtomicBool::new(false));
        let accepted = Arc::new(AtomicU64::new(0));
        let mut writer = StreamingDocumentWriter::new(
            CancellingSink {
                bytes: Vec::new(),
                source: source.clone(),
                armed: Arc::clone(&armed),
                cancelled: false,
                accepted: Arc::clone(&accepted),
            },
            context,
            limits(),
        )
        .expect("streaming writer");
        writer.start_paragraph().expect("paragraph");
        writer.finish_paragraph().expect("paragraph finish");
        armed.store(true, Ordering::Release);
        let error = match writer.finish() {
            Ok(_) => panic!("cancelling sink unexpectedly finished"),
            Err(error) => error,
        };
        assert!(matches!(error, StreamingDocumentError::Cancelled { .. }));
        assert_eq!(error.written(), accepted.load(Ordering::Acquire));
    }
}
