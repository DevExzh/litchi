//! Source-backed, content-only ODF publication.
//!
//! This module is deliberately narrower than the owning package rebuild path.
//! It publishes one already-authenticated `content.xml` replacement to a
//! sequential sink while retaining every other ZIP member's local span and
//! central-directory record through [`soapberry_zip::PreservationPlan`].
//! Unsupported source layouts are refused before the first sink write.  The
//! method does not materialize an owning package and has no logical fallback.

use super::package::{
    SourceBackedPackage, SourceChangedIo, SourceContentProof, SourceExecutionIo, SourceReadIo,
};
use crate::constants;
use litchi_core::{
    CancellationToken, Error, ExecutionContext, ExecutionError, Resource, Result, SourceVersion,
};
use soapberry_zip::{
    CompressionMethod, ErrorKind as ZipErrorKind, PreservationAction, PreservationPlan,
    RegeneratedEntry,
};
use std::collections::HashSet;
use std::fmt;
use std::io::{self, Write};
use std::sync::Arc;

const COPY_CHUNK_SIZE: usize = 32 * 1024;
const ZIP_LOCAL_HEADER_SIZE: usize = 30;
const ZIP_CENTRAL_HEADER_SIZE: usize = 46;
const ZIP_ENCRYPTED_FLAG: u16 = 1;
const ZIP_DATA_DESCRIPTOR_FLAG: u16 = 1 << 3;
const ZIP_UTF8_FLAG: u16 = 1 << 11;
const ZIP_STORE_METHOD: u16 = 0;
const DEFAULT_MAX_OUTPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;

/// Options for source-backed `content.xml` publication.
///
/// Both replacement and output ceilings are finite by default. The explicit
/// output ceiling is checked before each sink write and reports typed partial
/// progress when it is reached.
#[derive(Debug, Clone)]
pub struct SourceContentPublicationOptions {
    max_replacement_bytes: u64,
    max_output_bytes: u64,
    cancellation: Option<CancellationToken>,
    execution: Option<ExecutionContext>,
    verify_payloads: bool,
}

impl SourceContentPublicationOptions {
    /// Construct options with the common finite replacement ceiling.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_replacement_bytes: crate::package::MAX_CONTENT_REPLACEMENT_BYTES as u64,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
            cancellation: None,
            execution: None,
            verify_payloads: false,
        }
    }

    /// Set the maximum accepted replacement size.
    #[must_use]
    pub const fn with_max_replacement_bytes(mut self, maximum: u64) -> Self {
        self.max_replacement_bytes = maximum;
        self
    }

    /// Set the maximum complete output size accepted by the sink.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: u64) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Attach a cooperative cancellation token.
    #[must_use]
    pub fn with_cancellation(mut self, cancellation: CancellationToken) -> Self {
        self.cancellation = Some(cancellation);
        self
    }

    /// Attach a managed execution context. Its cancellation token and
    /// hierarchical budget are checked at bounded publication points.
    ///
    /// `InputBytes` charges exactly the bytes returned by positional source
    /// reads, `OutputBytes` charges only bytes accepted by the sink, `Work`
    /// charges source-copy, XML/compression, and requested payload verification
    /// work, `Objects` charges the bounded member plan, and `Memory` reserves a
    /// conservative publication envelope before any operation-owned allocation.
    #[must_use]
    pub fn with_execution_context(mut self, execution: ExecutionContext) -> Self {
        self.execution = Some(execution);
        self
    }

    /// Opt into reading and validating every untouched member before copying
    /// it. The default remains lazy for opaque media payloads.
    #[must_use]
    pub const fn with_payload_verification(mut self, verify: bool) -> Self {
        self.verify_payloads = verify;
        self
    }

    /// Return the replacement ceiling, if one is configured.
    #[must_use]
    pub const fn max_replacement_bytes(&self) -> u64 {
        self.max_replacement_bytes
    }

    /// Return the output ceiling, if one is configured.
    #[must_use]
    pub const fn max_output_bytes(&self) -> u64 {
        self.max_output_bytes
    }

    /// Return the cancellation token, if one is configured.
    #[must_use]
    pub const fn cancellation(&self) -> Option<&CancellationToken> {
        self.cancellation.as_ref()
    }

    /// Return the managed execution context, if one was attached.
    #[must_use]
    pub const fn execution_context(&self) -> Option<&ExecutionContext> {
        self.execution.as_ref()
    }

    /// Whether untouched payloads are verified before raw copying.
    #[must_use]
    pub const fn verify_payloads(&self) -> bool {
        self.verify_payloads
    }
}

impl Default for SourceContentPublicationOptions {
    fn default() -> Self {
        Self::new()
    }
}

/// Truthful progress for a non-seekable publication sink.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SourceContentPublicationProgress {
    /// No output byte was accepted.
    Untouched,
    /// The sink accepted a known prefix of the candidate artifact.
    Prefix {
        /// Bytes definitely accepted by the sink.
        accepted: u64,
    },
    /// Every candidate artifact byte was handed to the sink, but final flush
    /// did not complete.
    CompleteUnflushed {
        /// Bytes handed to the sink before flush failed.
        bytes: u64,
    },
    /// Every candidate artifact byte was handed to the sink and flush
    /// completed, but a later source/cancellation check failed.
    Complete {
        /// Bytes handed to the sink before the post-flush failure.
        bytes: u64,
    },
    /// A sink violated `Write` by reporting more bytes than supplied, so the
    /// exact accepted count after that call is unknowable.
    Indeterminate {
        /// Bytes definitely accepted before the invalid report.
        accepted_before: u64,
    },
}

impl SourceContentPublicationProgress {
    /// Return bytes definitely accepted before this progress state.
    #[must_use]
    pub const fn accepted(self) -> u64 {
        match self {
            Self::Untouched => 0,
            Self::Prefix { accepted }
            | Self::CompleteUnflushed { bytes: accepted }
            | Self::Complete { bytes: accepted }
            | Self::Indeterminate {
                accepted_before: accepted,
            } => accepted,
        }
    }
}

/// A source-backed content-only publication report.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceContentPublicationReport {
    bytes: u64,
    no_op: bool,
    source_version: SourceVersion,
}

impl SourceContentPublicationReport {
    /// Complete candidate artifact bytes accepted by the sink.
    #[must_use]
    pub const fn bytes(self) -> u64 {
        self.bytes
    }

    /// Whether the replacement was byte-identical to source `content.xml`.
    #[must_use]
    pub const fn is_no_op(self) -> bool {
        self.no_op
    }

    /// Source version checked before, during, and after publication.
    #[must_use]
    pub const fn source_version(self) -> SourceVersion {
        self.source_version
    }
}

/// Failure from a source-backed content-only publication.
#[derive(Debug)]
#[non_exhaustive]
pub enum SourceContentPublicationError {
    /// A normal ODF/source error occurred before output began.
    Core(Error),
    /// A bounded allocation failed before output began.
    Allocation {
        /// The bounded structure that could not be reserved.
        resource: &'static str,
        /// The allocator failure.
        source: std::collections::TryReserveError,
    },
    /// The source layout or policy is outside this exact raw-preservation
    /// contract.
    Unsupported {
        /// Why this source cannot be published through this operation.
        reason: String,
    },
    /// Cooperative cancellation was observed.
    Cancelled {
        /// Sink progress at cancellation.
        progress: SourceContentPublicationProgress,
    },
    /// A configured finite publication limit was reached.
    LimitExceeded {
        /// Sink progress at the limit failure.
        progress: SourceContentPublicationProgress,
        /// Attempted or observed value.
        actual: u64,
        /// Configured ceiling.
        maximum: u64,
    },
    /// The source version changed during publication.
    SourceChanged {
        /// Source version captured before publication.
        expected: SourceVersion,
        /// Version observed after or during publication.
        observed: SourceVersion,
        /// Sink progress at the stale-source failure.
        progress: SourceContentPublicationProgress,
    },
    /// The positional source failed for a reason other than a version change.
    Source {
        /// Sink progress when the source failure was observed.
        progress: SourceContentPublicationProgress,
        /// The original source error.
        source: Error,
    },
    /// An execution-context policy or resource budget rejected the operation.
    Execution {
        /// Sink progress when the execution failure was observed.
        progress: SourceContentPublicationProgress,
        /// The typed execution failure, including resource attribution.
        source: ExecutionError,
    },
    /// The caller-owned sink failed after or before accepting bytes.
    Sink {
        /// Sink progress at failure.
        progress: SourceContentPublicationProgress,
        /// Underlying sink error.
        source: io::Error,
    },
    /// The ZIP preservation layer rejected the source or write.
    Archive {
        /// Sink progress at failure.
        progress: SourceContentPublicationProgress,
        /// Underlying ZIP error.
        source: soapberry_zip::Error,
    },
}

impl SourceContentPublicationError {
    /// Return truthful sink progress for this failure.
    #[must_use]
    pub const fn progress(&self) -> SourceContentPublicationProgress {
        match self {
            Self::Core(_) | Self::Allocation { .. } | Self::Unsupported { .. } => {
                SourceContentPublicationProgress::Untouched
            },
            Self::Cancelled { progress }
            | Self::LimitExceeded { progress, .. }
            | Self::SourceChanged { progress, .. }
            | Self::Source { progress, .. }
            | Self::Execution { progress, .. }
            | Self::Sink { progress, .. }
            | Self::Archive { progress, .. } => *progress,
        }
    }

    /// Return bytes definitely accepted before failure.
    #[must_use]
    pub const fn written(&self) -> u64 {
        self.progress().accepted()
    }
}

impl fmt::Display for SourceContentPublicationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Core(error) => error.fmt(formatter),
            Self::Allocation { resource, source } => {
                write!(formatter, "could not reserve {resource}: {source}")
            },
            Self::Unsupported { reason } => {
                write!(
                    formatter,
                    "source-backed ODF content publication unavailable: {reason}"
                )
            },
            Self::Cancelled { progress } => {
                write!(
                    formatter,
                    "source-backed ODF content publication cancelled ({progress:?})"
                )
            },
            Self::LimitExceeded {
                progress,
                actual,
                maximum,
            } => write!(
                formatter,
                "source-backed ODF content publication limit exceeded ({actual} > {maximum}, {progress:?})"
            ),
            Self::SourceChanged {
                expected,
                observed,
                progress,
            } => write!(
                formatter,
                "source-backed ODF source changed from {expected:?} to {observed:?} ({progress:?})"
            ),
            Self::Source { progress, source } => {
                write!(
                    formatter,
                    "source-backed ODF source failed ({progress:?}): {source}"
                )
            },
            Self::Execution { progress, source } => write!(
                formatter,
                "source-backed ODF execution policy failed ({progress:?}): {source}"
            ),
            Self::Sink { progress, source } => {
                write!(
                    formatter,
                    "source-backed ODF sink failed ({progress:?}): {source}"
                )
            },
            Self::Archive { progress, source } => {
                write!(
                    formatter,
                    "source-backed ODF ZIP publication failed ({progress:?}): {source}"
                )
            },
        }
    }
}

impl std::error::Error for SourceContentPublicationError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Core(error) => Some(error),
            Self::Allocation { source, .. } => Some(source),
            Self::Source { source, .. } => Some(source),
            Self::Execution { source, .. } => Some(source),
            Self::Sink { source, .. } => Some(source),
            Self::Archive { source, .. } => Some(source),
            Self::Unsupported { .. }
            | Self::Cancelled { .. }
            | Self::LimitExceeded { .. }
            | Self::SourceChanged { .. } => None,
        }
    }
}

impl From<Error> for SourceContentPublicationError {
    fn from(error: Error) -> Self {
        Self::Core(error)
    }
}

/// Publish one source-backed ODF `content.xml` replacement to a non-seekable
/// sink using the default bounded policy.
pub fn write_content_xml_to_stream<W: Write>(
    package: &SourceBackedPackage,
    writer: W,
    replacement: impl AsRef<[u8]>,
) -> std::result::Result<SourceContentPublicationReport, SourceContentPublicationError> {
    write_content_xml_to_stream_with_options(
        package,
        writer,
        replacement,
        SourceContentPublicationOptions::new(),
    )
}

/// Publish one source-backed ODF `content.xml` replacement to a non-seekable
/// sink under explicit finite limits and cancellation.
pub fn write_content_xml_to_stream_with_options<W: Write>(
    package: &SourceBackedPackage,
    writer: W,
    replacement: impl AsRef<[u8]>,
    options: SourceContentPublicationOptions,
) -> std::result::Result<SourceContentPublicationReport, SourceContentPublicationError> {
    package.write_content_xml_to_stream_with_options(writer, replacement, options)
}

#[derive(Debug)]
enum PublicationIoFailure {
    Cancelled,
    Limit {
        actual: u64,
        maximum: u64,
    },
    Execution(ExecutionError),
    SourceChanged {
        expected: SourceVersion,
        observed: SourceVersion,
    },
}

impl fmt::Display for PublicationIoFailure {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Cancelled => formatter.write_str("source-backed publication was cancelled"),
            Self::Limit { actual, maximum } => {
                write!(
                    formatter,
                    "source-backed publication output limit {actual} > {maximum}"
                )
            },
            Self::Execution(error) => error.fmt(formatter),
            Self::SourceChanged { .. } => {
                formatter.write_str("source-backed publication source changed")
            },
        }
    }
}

impl std::error::Error for PublicationIoFailure {}

#[derive(Debug)]
struct SinkWriteIo {
    source: io::Error,
}

impl fmt::Display for SinkWriteIo {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.source.fmt(formatter)
    }
}

impl std::error::Error for SinkWriteIo {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(&self.source)
    }
}

struct OutputState {
    accepted: u64,
    overreported: bool,
    flush_started: bool,
    flush_completed: bool,
}

impl OutputState {
    fn new() -> Self {
        Self {
            accepted: 0,
            overreported: false,
            flush_started: false,
            flush_completed: false,
        }
    }

    fn progress(&self) -> SourceContentPublicationProgress {
        if self.overreported {
            SourceContentPublicationProgress::Indeterminate {
                accepted_before: self.accepted,
            }
        } else if self.flush_started && !self.flush_completed {
            SourceContentPublicationProgress::CompleteUnflushed {
                bytes: self.accepted,
            }
        } else if self.flush_completed {
            SourceContentPublicationProgress::Complete {
                bytes: self.accepted,
            }
        } else if self.accepted == 0 {
            SourceContentPublicationProgress::Untouched
        } else {
            SourceContentPublicationProgress::Prefix {
                accepted: self.accepted,
            }
        }
    }
}

struct CheckedSink<'a, W> {
    inner: W,
    package: &'a SourceBackedPackage,
    expected_version: SourceVersion,
    options: &'a SourceContentPublicationOptions,
    state: &'a mut OutputState,
}

impl<W: Write> CheckedSink<'_, W> {
    fn check_cancelled(&self) -> io::Result<()> {
        if self
            .options
            .cancellation
            .as_ref()
            .is_some_and(CancellationToken::is_cancelled)
        {
            return Err(io::Error::other(PublicationIoFailure::Cancelled));
        }
        if let Some(execution) = self.options.execution_context() {
            if let Err(error) = execution.check() {
                return Err(io::Error::other(publication_execution_failure(&error)));
            }
        }
        Ok(())
    }

    fn check_source(&self) -> io::Result<()> {
        let observed = self.package.source_version_observed().map_err(|source| {
            let kind = source.kind();
            io::Error::new(kind, SourceReadIo { source })
        })?;
        if observed != self.expected_version {
            return Err(io::Error::other(PublicationIoFailure::SourceChanged {
                expected: self.expected_version,
                observed,
            }));
        }
        Ok(())
    }
}

impl<W: Write> Write for CheckedSink<'_, W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.check_cancelled()?;
        self.check_source()?;
        let requested = u64::try_from(bytes.len()).map_err(|_| {
            io::Error::new(io::ErrorKind::InvalidInput, "publication write exceeds u64")
        })?;
        let maximum = self.options.max_output_bytes;
        let actual = self
            .state
            .accepted
            .checked_add(requested)
            .unwrap_or(u64::MAX);
        if actual > maximum {
            return Err(io::Error::other(PublicationIoFailure::Limit {
                actual,
                maximum,
            }));
        }

        let reservation = self
            .options
            .execution_context()
            .map(|execution| {
                execution
                    .reserve(Resource::OutputBytes, requested)
                    .map_err(|error| io::Error::other(publication_execution_failure(&error)))
            })
            .transpose()?;

        let sink_result = {
            let _input_suspension = self.package.suspend_publication_input_accounting();
            self.inner.write(bytes)
        };
        let written = match sink_result {
            Ok(written) => written,
            Err(error) => {
                self.state.overreported = true;
                let kind = error.kind();
                return Err(io::Error::new(kind, SinkWriteIo { source: error }));
            },
        };
        if written == 0 && !bytes.is_empty() {
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                SinkWriteIo {
                    source: io::Error::new(
                        io::ErrorKind::WriteZero,
                        "publication sink accepted no bytes",
                    ),
                },
            ));
        }
        if written > bytes.len() {
            self.state.overreported = true;
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "publication sink reported more bytes than supplied",
            ));
        }
        let written_u64 = u64::try_from(written).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                "publication progress exceeds u64",
            )
        })?;
        self.state.accepted = self
            .state
            .accepted
            .checked_add(written_u64)
            .ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidData, "publication progress overflow")
            })?;
        if let Some(reservation) = reservation {
            let _ = reservation.commit(written_u64);
        }
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.state.flush_started = true;
        self.check_cancelled()?;
        self.check_source()?;
        {
            let _input_suspension = self.package.suspend_publication_input_accounting();
            self.inner.flush()?;
        }
        self.state.flush_completed = true;
        self.check_source()?;
        self.check_cancelled()
    }
}

impl SourceBackedPackage {
    /// Publish one `content.xml` replacement to a sequential sink.
    ///
    /// The method is intentionally content-only. Exact byte-identical
    /// replacements copy the complete source artifact byte-for-byte, including
    /// signed or otherwise unsupported unencrypted source details. Encrypted
    /// sources are refused before content materialization. A changed
    /// replacement is accepted only when every preflight proof required by the
    /// raw-preservation contract succeeds; there is no owning-package fallback.
    /// Untouched members with compression methods unknown to Litchi remain
    /// byte-preserved but opaque, so this operation does not promise that those
    /// members can later be decompressed by Litchi.
    pub fn write_content_xml_to_stream<W: Write>(
        &self,
        writer: W,
        replacement: impl AsRef<[u8]>,
    ) -> std::result::Result<SourceContentPublicationReport, SourceContentPublicationError> {
        self.write_content_xml_to_stream_with_options(
            writer,
            replacement,
            SourceContentPublicationOptions::new(),
        )
    }

    /// Publish one `content.xml` replacement under explicit finite limits and
    /// cancellation.
    pub fn write_content_xml_to_stream_with_options<W: Write>(
        &self,
        writer: W,
        replacement: impl AsRef<[u8]>,
        options: SourceContentPublicationOptions,
    ) -> std::result::Result<SourceContentPublicationReport, SourceContentPublicationError> {
        let replacement_bytes = replacement.as_ref();
        let result =
            self.write_content_xml_to_stream_inner(writer, replacement_bytes, None, &options);
        match result {
            Ok(report) => Ok(report),
            Err(error) => match reconcile_source_state(self, error.progress()) {
                Ok(()) => Err(error),
                Err(source_error) => Err(source_error),
            },
        }
    }

    /// Publish using `content.xml` bytes already authenticated by this exact
    /// source owner.  This is a hidden facade bridge: callers must retain the
    /// [`SourceContentProof`] returned by
    /// [`SourceBackedPackage::get_content_xml_with_source_proof`] and pass the
    /// matching bytes unchanged.  All normal publication gates still run.
    #[doc(hidden)]
    pub fn write_content_xml_to_stream_with_known_source_content<W: Write>(
        &self,
        writer: W,
        known_content: &[u8],
        proof: &SourceContentProof,
        replacement: impl AsRef<[u8]>,
        options: SourceContentPublicationOptions,
    ) -> std::result::Result<SourceContentPublicationReport, SourceContentPublicationError> {
        let replacement_bytes = replacement.as_ref();
        let result = self.write_content_xml_to_stream_inner(
            writer,
            replacement_bytes,
            Some((proof, known_content)),
            &options,
        );
        match result {
            Ok(report) => Ok(report),
            Err(error) => match reconcile_source_state(self, error.progress()) {
                Ok(()) => Err(error),
                Err(source_error) => Err(source_error),
            },
        }
    }

    fn write_content_xml_to_stream_inner<W: Write>(
        &self,
        writer: W,
        replacement_bytes: &[u8],
        known_source_content: Option<(&SourceContentProof, &[u8])>,
        options: &SourceContentPublicationOptions,
    ) -> std::result::Result<SourceContentPublicationReport, SourceContentPublicationError> {
        check_cancellation(options, SourceContentPublicationProgress::Untouched)?;
        let replacement_len = u64::try_from(replacement_bytes.len()).unwrap_or(u64::MAX);
        validate_replacement_limit(replacement_len, options)?;
        self.ensure_current_for_publication()
            .map_err(|error| map_core_error(error, SourceContentPublicationProgress::Untouched))?;
        reject_encrypted_source(self)?;

        let content_size = self
            .publication_member_size(constants::ODF_CONTENT)
            .map_err(|error| map_core_error(error, SourceContentPublicationProgress::Untouched))?;
        let entry_count = u64::try_from(self.publication_entry_count()).unwrap_or(u64::MAX);
        let metadata_bytes = self.publication_metadata_bytes();
        let maximum_member = if options.verify_payloads() {
            self.publication_max_member_size().map_err(|error| {
                map_core_error(error, SourceContentPublicationProgress::Untouched)
            })?
        } else {
            content_size
        };
        let memory_required = publication_memory_requirement(
            replacement_len,
            maximum_member,
            metadata_bytes,
            entry_count,
        )?;
        let _operation_memory = reserve_memory(
            options,
            memory_required,
            SourceContentPublicationProgress::Untouched,
        )?;
        let _input_accounting = self
            .begin_publication_input_accounting(options.execution_context())
            .map_err(|source| SourceContentPublicationError::Allocation {
                resource: "source-backed ODF publication input-accounting stack",
                source,
            })?;
        consume_execution_resource(
            options,
            Resource::Work,
            content_size,
            SourceContentPublicationProgress::Untouched,
        )?;
        consume_execution_resource(
            options,
            Resource::Objects,
            1,
            SourceContentPublicationProgress::Untouched,
        )?;

        let no_op = if let Some((proof, known_content)) = known_source_content {
            self.validate_known_content_proof(proof, known_content)
                .map_err(|error| {
                    map_core_error(error, SourceContentPublicationProgress::Untouched)
                })?;
            known_content == replacement_bytes
        } else {
            let source_content = self.get_file(constants::ODF_CONTENT).map_err(|error| {
                map_core_error(error, SourceContentPublicationProgress::Untouched)
            })?;
            self.ensure_current_for_publication().map_err(|error| {
                map_core_error(error, SourceContentPublicationProgress::Untouched)
            })?;
            let no_op = source_content == replacement_bytes;
            drop(source_content);
            no_op
        };

        if no_op {
            consume_execution_resource(
                options,
                Resource::Work,
                self.len(),
                SourceContentPublicationProgress::Untouched,
            )?;
            if self.len() > options.max_output_bytes {
                return Err(SourceContentPublicationError::LimitExceeded {
                    progress: SourceContentPublicationProgress::Untouched,
                    actual: self.len(),
                    maximum: options.max_output_bytes,
                });
            }
            let mut state = OutputState::new();
            let result = write_exact_source(self, writer, options, &mut state);
            return finish_publication(self, options, true, state, result);
        }

        consume_execution_resource(
            options,
            Resource::Work,
            replacement_len
                .checked_mul(2)
                .and_then(|replacement_work| self.len().checked_add(replacement_work))
                .ok_or_else(|| unsupported("source-backed ODF work accounting overflow"))?,
            SourceContentPublicationProgress::Untouched,
        )?;
        consume_execution_resource(
            options,
            Resource::Objects,
            entry_count
                .checked_add(3)
                .ok_or_else(|| unsupported("source-backed ODF object accounting overflow"))?,
            SourceContentPublicationProgress::Untouched,
        )?;
        let replacement = clone_replacement(replacement_bytes)?;
        preflight_changed_publication(self, &replacement, options)?;
        let mut scratch = Vec::new();
        scratch
            .try_reserve_exact(soapberry_zip::RECOMMENDED_BUFFER_SIZE)
            .map_err(|source| SourceContentPublicationError::Allocation {
                resource: "source-backed ODF preservation index",
                source,
            })?;
        scratch.resize(soapberry_zip::RECOMMENDED_BUFFER_SIZE, 0);
        let index = self
            .preservation_index(&mut scratch)
            .map_err(map_preflight_zip_error)?;
        self.ensure_current_for_publication()
            .map_err(|error| map_core_error(error, SourceContentPublicationProgress::Untouched))?;
        if index.archive_end_offset() != self.len() {
            return Err(unsupported(
                "source archive has trailing bytes outside the located ZIP",
            ));
        }
        validate_canonical_mimetype(self, &index)?;
        validate_preserved_member_framing(self, &index)?;
        if options.verify_payloads() {
            verify_untouched_payloads(self, options)?;
        }
        let replacement = Arc::new(replacement);
        let mut plan = PreservationPlan::new();
        plan.try_reserve_exact(index.entries().len())
            .map_err(|source| SourceContentPublicationError::Allocation {
                resource: "source-backed ODF preservation actions",
                source,
            })?;
        let mut replaced = false;
        let mut seen_names = HashSet::new();
        seen_names
            .try_reserve(index.entries().len())
            .map_err(|source| SourceContentPublicationError::Allocation {
                resource: "source-backed ODF preservation member names",
                source,
            })?;
        for entry in index.entries() {
            let raw_name = entry.raw_name_bytes();
            let mut owned_name = Vec::new();
            owned_name
                .try_reserve_exact(raw_name.len())
                .map_err(|source| SourceContentPublicationError::Allocation {
                    resource: "source-backed ODF preservation member name",
                    source,
                })?;
            owned_name.extend_from_slice(raw_name);
            if !seen_names.insert(owned_name) {
                return Err(unsupported("ambiguous duplicate ZIP member names"));
            }
            if raw_name == constants::ODF_CONTENT.as_bytes() {
                if replaced {
                    return Err(unsupported("ambiguous duplicate content.xml members"));
                }
                let compression = match entry.compression_method() {
                    CompressionMethod::Store => CompressionMethod::Store,
                    CompressionMethod::Deflate => CompressionMethod::Deflate,
                    _ => {
                        return Err(unsupported("content.xml uses unsupported ZIP compression"));
                    },
                };
                plan.push(PreservationAction::Regenerate {
                    id: entry.id(),
                    entry: RegeneratedEntry::new_shared(
                        constants::ODF_CONTENT,
                        Arc::clone(&replacement),
                    )
                    .compression_method(compression),
                });
                replaced = true;
            } else {
                plan.push(PreservationAction::Copy(entry.id()));
            }
        }
        if !replaced {
            return Err(unsupported(
                "source archive has no canonical content.xml member",
            ));
        }

        check_cancellation(options, SourceContentPublicationProgress::Untouched)?;
        let mut state = OutputState::new();
        let expected_version = self.source_version_snapshot();
        let checked = CheckedSink {
            inner: writer,
            package: self,
            expected_version,
            options,
            state: &mut state,
        };
        let result = match index.write_to(&plan, checked) {
            Ok(_sink) => Ok(()),
            Err(error) => Err(map_zip_error(error, state.progress())),
        };
        finish_publication(self, options, false, state, result)
    }
}

fn clone_replacement(bytes: &[u8]) -> std::result::Result<Vec<u8>, SourceContentPublicationError> {
    let mut replacement = Vec::new();
    replacement
        .try_reserve_exact(bytes.len())
        .map_err(|source| SourceContentPublicationError::Allocation {
            resource: "source-backed ODF content.xml replacement",
            source,
        })?;
    replacement.extend_from_slice(bytes);
    Ok(replacement)
}

fn validate_replacement_limit(
    observed: u64,
    options: &SourceContentPublicationOptions,
) -> std::result::Result<(), SourceContentPublicationError> {
    let maximum = options.max_replacement_bytes;
    if observed > maximum {
        return Err(SourceContentPublicationError::LimitExceeded {
            progress: SourceContentPublicationProgress::Untouched,
            actual: observed,
            maximum,
        });
    }
    Ok(())
}

fn reject_encrypted_source(
    package: &SourceBackedPackage,
) -> std::result::Result<(), SourceContentPublicationError> {
    let manifest = package
        .manifest()
        .map_err(|error| map_core_error(error, SourceContentPublicationProgress::Untouched))?;
    if manifest.has_encrypted_entries() || package.publication_has_zip_encrypted_entries() {
        return Err(unsupported(
            "encrypted ODF entries cannot be published without materialization",
        ));
    }
    Ok(())
}

fn preflight_changed_publication(
    package: &SourceBackedPackage,
    replacement: &[u8],
    options: &SourceContentPublicationOptions,
) -> std::result::Result<(), SourceContentPublicationError> {
    check_cancellation(options, SourceContentPublicationProgress::Untouched)?;
    let manifest = package
        .manifest()
        .map_err(|error| map_core_error(error, SourceContentPublicationProgress::Untouched))?;
    let Some(content_entry) = manifest.get_entry(constants::ODF_CONTENT) else {
        return Err(unsupported("manifest has no canonical content.xml entry"));
    };
    if content_entry.size.is_some() {
        return Err(unsupported(
            "manifest:size on content.xml is incompatible with raw replacement",
        ));
    }
    if !xml_minifier::audit::package::is_xml_part(constants::ODF_CONTENT, &content_entry.media_type)
    {
        return Err(unsupported("content.xml manifest media type is not XML"));
    }
    debug_assert!(!manifest.has_encrypted_entries());
    for path in package.publication_file_names() {
        if path.starts_with("META-INF/") && path.ends_with("signatures.xml") {
            return Err(unsupported(
                "signed ODF packages require an explicit signature policy",
            ));
        }
    }
    let _ =
        xml_minifier::audit::verify_authored(replacement, xml_minifier::audit::Limits::default())
            .map_err(|error| {
            SourceContentPublicationError::Core(Error::InvalidFormat(format!(
                "XML publication rejected for 'content.xml': {error}"
            )))
        })?;
    package
        .ensure_current_for_publication()
        .map_err(|error| map_core_error(error, SourceContentPublicationProgress::Untouched))?;
    Ok(())
}

fn validate_preserved_member_framing(
    package: &SourceBackedPackage,
    index: &soapberry_zip::PreservationIndex<'_, impl soapberry_zip::ReaderAt>,
) -> std::result::Result<(), SourceContentPublicationError> {
    for entry in index.entries() {
        let local_span = entry.local_span();
        let central_span = entry.central_record();
        let mut local = [0_u8; ZIP_LOCAL_HEADER_SIZE];
        let mut central = [0_u8; ZIP_CENTRAL_HEADER_SIZE];
        read_source_exact(package, local_span.start, &mut local).map_err(|error| {
            map_source_error(error, SourceContentPublicationProgress::Untouched)
        })?;
        read_source_exact(package, central_span.start, &mut central).map_err(|error| {
            map_source_error(error, SourceContentPublicationProgress::Untouched)
        })?;

        let local_name_len = le_u16(&local, 26).map(u64::from);
        let local_extra_len = le_u16(&local, 28).map(u64::from);
        let local_name_size = usize::try_from(local_name_len.unwrap_or(u64::MAX))
            .map_err(|_| unsupported("local ZIP member name exceeds platform limits"))?;
        if local_name_size != entry.raw_name_bytes().len() {
            return Err(unsupported(
                "local and central ZIP member-name lengths differ",
            ));
        }
        let mut local_name = Vec::new();
        local_name
            .try_reserve_exact(local_name_size)
            .map_err(|source| SourceContentPublicationError::Allocation {
                resource: "source-backed ODF local ZIP member name",
                source,
            })?;
        local_name.resize(local_name_size, 0);
        let local_name_offset = local_span
            .start
            .checked_add(ZIP_LOCAL_HEADER_SIZE as u64)
            .ok_or_else(|| unsupported("local ZIP member-name range overflow"))?;
        read_source_exact(package, local_name_offset, &mut local_name).map_err(|error| {
            map_source_error(error, SourceContentPublicationProgress::Untouched)
        })?;
        if local_name != entry.raw_name_bytes() {
            return Err(unsupported("local and central ZIP member names differ"));
        }
        let payload_start = local_span
            .start
            .checked_add(ZIP_LOCAL_HEADER_SIZE as u64)
            .and_then(|offset| offset.checked_add(local_name_len.unwrap_or(u64::MAX)))
            .and_then(|offset| offset.checked_add(local_extra_len.unwrap_or(u64::MAX)))
            .ok_or_else(|| unsupported("local ZIP payload range overflow"))?;
        let compressed_size = le_u32(&central, 20)
            .map(u64::from)
            .ok_or_else(|| unsupported("truncated central ZIP compressed-size field"))?;
        let uncompressed_size = le_u32(&central, 24)
            .map(u64::from)
            .ok_or_else(|| unsupported("truncated central ZIP uncompressed-size field"))?;
        let payload_end = payload_start
            .checked_add(compressed_size)
            .ok_or_else(|| unsupported("local ZIP payload range overflow"))?;
        if payload_end > local_span.end {
            return Err(unsupported("truncated preserved ZIP member payload"));
        }

        let flags =
            le_u16(&central, 8).ok_or_else(|| unsupported("truncated central ZIP flags"))?;
        if flags & ZIP_ENCRYPTED_FLAG != 0 {
            return Err(unsupported(
                "ZIP-encrypted ODF entries cannot be raw-copied",
            ));
        }
        if flags & ZIP_DATA_DESCRIPTOR_FLAG == 0 {
            if payload_end != local_span.end {
                return Err(unsupported(
                    "non-descriptor ZIP member has opaque trailing bytes",
                ));
            }
            continue;
        }

        let descriptor_len = usize::try_from(local_span.end - payload_end)
            .map_err(|_| unsupported("ZIP data descriptor length exceeds platform limits"))?;
        if !matches!(descriptor_len, 12 | 16 | 20 | 24) {
            return Err(unsupported("ZIP data descriptor has noncanonical framing"));
        }
        let mut descriptor = [0_u8; 24];
        read_source_exact(package, payload_end, &mut descriptor[..descriptor_len]).map_err(
            |error| map_source_error(error, SourceContentPublicationProgress::Untouched),
        )?;
        let descriptor = &descriptor[..descriptor_len];
        let expected_crc =
            le_u32(&central, 16).ok_or_else(|| unsupported("truncated central ZIP CRC"))?;
        let (descriptor_crc, descriptor_compressed, descriptor_uncompressed) = match descriptor_len
        {
            12 if le_u32(descriptor, 0) != Some(0x0807_4b50) => (
                le_u32(descriptor, 0),
                le_u32(descriptor, 4).map(u64::from),
                le_u32(descriptor, 8).map(u64::from),
            ),
            16 if le_u32(descriptor, 0) == Some(0x0807_4b50) => (
                le_u32(descriptor, 4),
                le_u32(descriptor, 8).map(u64::from),
                le_u32(descriptor, 12).map(u64::from),
            ),
            20 if le_u32(descriptor, 0) != Some(0x0807_4b50) => (
                le_u32(descriptor, 0),
                le_u64(descriptor, 4),
                le_u64(descriptor, 12),
            ),
            24 if le_u32(descriptor, 0) == Some(0x0807_4b50) => (
                le_u32(descriptor, 4),
                le_u64(descriptor, 8),
                le_u64(descriptor, 16),
            ),
            _ => return Err(unsupported("ZIP data descriptor signature is invalid")),
        };
        let local_crc = le_u32(&local, 14);
        let local_compressed = le_u32(&local, 18).map(u64::from);
        let local_uncompressed = le_u32(&local, 22).map(u64::from);
        if descriptor_crc != Some(expected_crc)
            || descriptor_compressed != Some(compressed_size)
            || descriptor_uncompressed != Some(uncompressed_size)
            || !((local_crc == Some(0)
                && local_compressed == Some(0)
                && local_uncompressed == Some(0))
                || (local_crc == Some(expected_crc)
                    && local_compressed == Some(compressed_size)
                    && local_uncompressed == Some(uncompressed_size)))
        {
            return Err(unsupported("ZIP data descriptor metadata is inconsistent"));
        }
    }
    Ok(())
}

fn validate_canonical_mimetype(
    package: &SourceBackedPackage,
    index: &soapberry_zip::PreservationIndex<'_, impl soapberry_zip::ReaderAt>,
) -> std::result::Result<(), SourceContentPublicationError> {
    if !constants::is_odf_mime_type(
        package
            .mimetype()
            .map_err(|error| map_core_error(error, SourceContentPublicationProgress::Untouched))?,
    ) || package
        .manifest()
        .map_err(|error| map_core_error(error, SourceContentPublicationProgress::Untouched))?
        .mimetype
        != package
            .mimetype()
            .map_err(|error| map_core_error(error, SourceContentPublicationProgress::Untouched))?
    {
        return Err(unsupported("noncanonical ODF mimetype metadata"));
    }
    let mut mimetype_entries = index
        .entries()
        .iter()
        .filter(|candidate| candidate.raw_name_bytes() == b"mimetype");
    let Some(entry) = mimetype_entries.next() else {
        return Err(unsupported(
            "source archive must contain exactly one mimetype member",
        ));
    };
    if mimetype_entries.next().is_some() {
        return Err(unsupported(
            "source archive must contain exactly one mimetype member",
        ));
    }
    if entry.local_span().start != 0 || entry.compression_method() != CompressionMethod::Store {
        return Err(unsupported(
            "mimetype is not the canonical first stored member",
        ));
    }

    let mut local = [0_u8; ZIP_LOCAL_HEADER_SIZE];
    read_source_exact(package, 0, &mut local)
        .map_err(|error| map_source_error(error, SourceContentPublicationProgress::Untouched))?;
    let local_name_len = le_u16(&local, 26).map(usize::from);
    let local_extra_len = le_u16(&local, 28).map(usize::from);
    let local_compressed = le_u32(&local, 18).map(u64::from);
    let local_uncompressed = le_u32(&local, 22).map(u64::from);
    if le_u32(&local, 0) != Some(0x0403_4b50)
        || le_u16(&local, 6).is_none_or(|flags| flags & !ZIP_UTF8_FLAG != 0)
        || le_u16(&local, 8) != Some(ZIP_STORE_METHOD)
        || local_name_len != Some(b"mimetype".len())
        || local_extra_len != Some(0)
        || local_compressed != local_uncompressed
    {
        return Err(unsupported("mimetype local header is not canonical"));
    }
    let payload_len = usize::try_from(local_compressed.unwrap_or(0))
        .map_err(|_| unsupported("mimetype payload exceeds platform limits"))?;
    let payload_start = ZIP_LOCAL_HEADER_SIZE + b"mimetype".len();
    let payload_end = payload_start
        .checked_add(payload_len)
        .ok_or_else(|| unsupported("mimetype payload range overflow"))?;
    if payload_end as u64 != entry.local_span().end
        || payload_len
            != package
                .mimetype()
                .map_err(|error| {
                    map_core_error(error, SourceContentPublicationProgress::Untouched)
                })?
                .len()
    {
        return Err(unsupported(
            "mimetype local payload framing is not canonical",
        ));
    }
    let mut payload = Vec::new();
    payload.try_reserve_exact(payload_len).map_err(|source| {
        SourceContentPublicationError::Allocation {
            resource: "source-backed ODF mimetype verification",
            source,
        }
    })?;
    payload.resize(payload_len, 0);
    read_source_exact(package, payload_start as u64, &mut payload)
        .map_err(|error| map_source_error(error, SourceContentPublicationProgress::Untouched))?;
    let mimetype = package
        .mimetype()
        .map_err(|error| map_core_error(error, SourceContentPublicationProgress::Untouched))?;
    if payload != mimetype.as_bytes() || le_u32(&local, 14) != Some(soapberry_zip::crc32(&payload))
    {
        return Err(unsupported("mimetype payload is not canonical"));
    }

    let central = entry.central_record();
    let mut fixed = [0_u8; ZIP_CENTRAL_HEADER_SIZE];
    read_source_exact(package, central.start, &mut fixed)
        .map_err(|error| map_source_error(error, SourceContentPublicationProgress::Untouched))?;
    let name_len = le_u16(&fixed, 28).map(usize::from);
    let extra_len = le_u16(&fixed, 30).map(usize::from);
    let comment_len = le_u16(&fixed, 32).map(usize::from);
    let central_len = ZIP_CENTRAL_HEADER_SIZE
        .checked_add(name_len.unwrap_or(usize::MAX))
        .and_then(|length| length.checked_add(extra_len.unwrap_or(usize::MAX)))
        .and_then(|length| length.checked_add(comment_len.unwrap_or(usize::MAX)));
    if le_u32(&fixed, 0) != Some(0x0201_4b50)
        || name_len != Some(b"mimetype".len())
        || extra_len != Some(0)
        || central_len != Some((central.end - central.start) as usize)
        || le_u16(&fixed, 8) != le_u16(&local, 6)
        || le_u16(&fixed, 10) != Some(ZIP_STORE_METHOD)
        || le_u32(&fixed, 16) != le_u32(&local, 14)
        || le_u32(&fixed, 20) != le_u32(&local, 18)
        || le_u32(&fixed, 24) != le_u32(&local, 22)
    {
        return Err(unsupported("mimetype central header is not canonical"));
    }
    Ok(())
}

fn verify_untouched_payloads(
    package: &SourceBackedPackage,
    options: &SourceContentPublicationOptions,
) -> std::result::Result<(), SourceContentPublicationError> {
    for path in package.publication_file_names() {
        check_cancellation(options, SourceContentPublicationProgress::Untouched)?;
        if path.ends_with('/') || path == constants::ODF_CONTENT {
            continue;
        }
        let payload_len = package
            .publication_member_size(path)
            .map_err(|error| map_core_error(error, SourceContentPublicationProgress::Untouched))?;
        consume_execution_resource(
            options,
            Resource::Work,
            payload_len,
            SourceContentPublicationProgress::Untouched,
        )?;
        let payload = package
            .get_file(path)
            .map_err(|error| map_core_error(error, SourceContentPublicationProgress::Untouched))?;
        if u64::try_from(payload.len()).unwrap_or(u64::MAX) != payload_len {
            return Err(unsupported(
                "verified ODF payload length differs from declared metadata",
            ));
        }
    }
    Ok(())
}

fn read_source_exact(package: &SourceBackedPackage, offset: u64, output: &mut [u8]) -> Result<()> {
    package.ensure_current_for_publication()?;
    match package.source_read_exact_at(offset, output) {
        Ok(()) => package.ensure_current_for_publication(),
        Err(error) => {
            package.ensure_current_for_publication()?;
            Err(Error::Io(error))
        },
    }
}

fn write_exact_source<W: Write>(
    package: &SourceBackedPackage,
    writer: W,
    options: &SourceContentPublicationOptions,
    state: &mut OutputState,
) -> std::result::Result<(), SourceContentPublicationError> {
    let expected_version = package.source_version_snapshot();
    let mut sink = CheckedSink {
        inner: writer,
        package,
        expected_version,
        options,
        state,
    };
    let mut buffer = [0_u8; COPY_CHUNK_SIZE];
    let mut offset = 0_u64;
    while offset < package.len() {
        let progress = sink.state.progress();
        check_cancellation(options, progress)?;
        package
            .ensure_current_for_publication()
            .map_err(|error| map_core_error(error, progress))?;
        let remaining = package.len() - offset;
        let requested = usize::try_from(remaining.min(buffer.len() as u64))
            .map_err(|_| unsupported("source copy range does not fit this platform"))?;
        let read = match package.source_read_at(offset, &mut buffer[..requested]) {
            Ok(read) => read,
            Err(error) => {
                package
                    .ensure_current_for_publication()
                    .map_err(|changed| map_core_error(changed, progress))?;
                return Err(map_owned_io_error(error, progress, UnmarkedIo::Source));
            },
        };
        package
            .ensure_current_for_publication()
            .map_err(|error| map_core_error(error, progress))?;
        if read == 0 {
            return Err(SourceContentPublicationError::Source {
                progress,
                source: Error::Io(io::Error::new(
                    io::ErrorKind::UnexpectedEof,
                    "source ended during exact ODF publication",
                )),
            });
        }
        if read > requested {
            return Err(SourceContentPublicationError::Source {
                progress,
                source: Error::Io(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "source reported more bytes than requested during exact publication",
                )),
            });
        }
        sink.write_all(&buffer[..read])
            .map_err(|error| map_io_error(error, sink.state.progress()))?;
        offset = offset.checked_add(read as u64).ok_or_else(|| {
            SourceContentPublicationError::Source {
                progress: sink.state.progress(),
                source: Error::InvalidFormat("source copy offset overflow".to_string()),
            }
        })?;
    }
    sink.flush()
        .map_err(|error| map_io_error(error, sink.state.progress()))
}

fn finish_publication(
    package: &SourceBackedPackage,
    options: &SourceContentPublicationOptions,
    no_op: bool,
    state: OutputState,
    result: std::result::Result<(), SourceContentPublicationError>,
) -> std::result::Result<SourceContentPublicationReport, SourceContentPublicationError> {
    result?;
    let progress = state.progress();
    check_cancellation(options, progress)?;
    reconcile_source_state(package, progress)?;
    Ok(SourceContentPublicationReport {
        bytes: progress.accepted(),
        no_op,
        source_version: package.source_version_snapshot(),
    })
}

fn reconcile_source_state(
    package: &SourceBackedPackage,
    progress: SourceContentPublicationProgress,
) -> std::result::Result<(), SourceContentPublicationError> {
    let observed_version = package.source_version_observed().map_err(|error| {
        SourceContentPublicationError::Source {
            progress,
            source: Error::Io(error),
        }
    })?;
    if observed_version != package.source_version_snapshot() {
        return Err(SourceContentPublicationError::SourceChanged {
            expected: package.source_version_snapshot(),
            observed: observed_version,
            progress,
        });
    }
    let observed_length = package.source_length_observed().map_err(|error| {
        SourceContentPublicationError::Source {
            progress,
            source: Error::Io(error),
        }
    })?;
    let final_version = package.source_version_observed().map_err(|error| {
        SourceContentPublicationError::Source {
            progress,
            source: Error::Io(error),
        }
    })?;
    if final_version != package.source_version_snapshot() {
        return Err(SourceContentPublicationError::SourceChanged {
            expected: package.source_version_snapshot(),
            observed: final_version,
            progress,
        });
    }
    if observed_length != package.len() {
        return Err(SourceContentPublicationError::Source {
            progress,
            source: Error::InvalidFormat(
                "source length changed without a new source version".to_string(),
            ),
        });
    }
    Ok(())
}

fn check_cancellation(
    options: &SourceContentPublicationOptions,
    progress: SourceContentPublicationProgress,
) -> std::result::Result<(), SourceContentPublicationError> {
    if options
        .cancellation
        .as_ref()
        .is_some_and(CancellationToken::is_cancelled)
    {
        return Err(SourceContentPublicationError::Cancelled { progress });
    }
    if let Some(execution) = options.execution.as_ref() {
        execution
            .check()
            .map_err(|error| map_execution_error(error, progress))?;
    }
    Ok(())
}

fn consume_execution_resource(
    options: &SourceContentPublicationOptions,
    resource: Resource,
    amount: u64,
    progress: SourceContentPublicationProgress,
) -> std::result::Result<(), SourceContentPublicationError> {
    check_cancellation(options, progress)?;
    if let Some(execution) = options.execution.as_ref() {
        execution
            .consume(resource, amount)
            .map_err(|error| map_execution_error(error, progress))?;
    }
    Ok(())
}

fn publication_memory_requirement(
    replacement_len: u64,
    maximum_member: u64,
    metadata_bytes: u64,
    entry_count: u64,
) -> std::result::Result<u64, SourceContentPublicationError> {
    replacement_len
        .checked_mul(3)
        .and_then(|amount| amount.checked_add(maximum_member))
        .and_then(|amount| {
            metadata_bytes
                .checked_mul(2)
                .and_then(|metadata| amount.checked_add(metadata))
        })
        .and_then(|amount| {
            amount.checked_add(
                u64::try_from(soapberry_zip::RECOMMENDED_BUFFER_SIZE + COPY_CHUNK_SIZE)
                    .unwrap_or(u64::MAX),
            )
        })
        .and_then(|amount| {
            entry_count
                .checked_mul(512)
                .and_then(|metadata| amount.checked_add(metadata))
        })
        .ok_or_else(|| unsupported("source-backed ODF memory accounting overflow"))
}

fn reserve_memory(
    options: &SourceContentPublicationOptions,
    amount: u64,
    progress: SourceContentPublicationProgress,
) -> std::result::Result<Option<litchi_core::Reservation>, SourceContentPublicationError> {
    let Some(execution) = options.execution_context() else {
        return Ok(None);
    };
    execution
        .reserve(Resource::Memory, amount)
        .map(Some)
        .map_err(|error| map_execution_error(error, progress))
}

fn map_core_error(
    error: Error,
    progress: SourceContentPublicationProgress,
) -> SourceContentPublicationError {
    match error {
        Error::SourceChanged { expected, observed } => {
            SourceContentPublicationError::SourceChanged {
                expected,
                observed,
                progress,
            }
        },
        Error::Allocation { resource, source } => {
            SourceContentPublicationError::Allocation { resource, source }
        },
        Error::Io(source) => map_owned_io_error(source, progress, UnmarkedIo::Source),
        other => SourceContentPublicationError::Core(other),
    }
}

fn map_source_error(
    error: Error,
    progress: SourceContentPublicationProgress,
) -> SourceContentPublicationError {
    match error {
        Error::SourceChanged { expected, observed } => {
            SourceContentPublicationError::SourceChanged {
                expected,
                observed,
                progress,
            }
        },
        Error::Io(source) => map_owned_io_error(source, progress, UnmarkedIo::Source),
        source => SourceContentPublicationError::Source { progress, source },
    }
}

fn map_execution_error(
    error: ExecutionError,
    progress: SourceContentPublicationProgress,
) -> SourceContentPublicationError {
    match error {
        ExecutionError::Cancelled => SourceContentPublicationError::Cancelled { progress },
        source => SourceContentPublicationError::Execution { progress, source },
    }
}

fn map_preflight_zip_error(error: soapberry_zip::Error) -> SourceContentPublicationError {
    match error.into_kind() {
        ZipErrorKind::Allocation { resource, source } => {
            SourceContentPublicationError::Allocation { resource, source }
        },
        ZipErrorKind::UnsupportedPreservation { reason } => unsupported(reason),
        ZipErrorKind::UnsupportedCompressionMethod(method) => {
            unsupported(format!("unsupported source compression method {method}"))
        },
        ZipErrorKind::IO(source) | ZipErrorKind::Io(source) => map_owned_io_error(
            source,
            SourceContentPublicationProgress::Untouched,
            UnmarkedIo::Source,
        ),
        kind => SourceContentPublicationError::Archive {
            progress: SourceContentPublicationProgress::Untouched,
            source: soapberry_zip::Error::from(kind),
        },
    }
}

fn map_zip_error(
    error: soapberry_zip::Error,
    progress: SourceContentPublicationProgress,
) -> SourceContentPublicationError {
    match error.into_kind() {
        ZipErrorKind::Allocation { resource, source } => {
            SourceContentPublicationError::Allocation { resource, source }
        },
        ZipErrorKind::UnsupportedPreservation { reason } => {
            SourceContentPublicationError::Unsupported {
                reason: reason.to_string(),
            }
        },
        ZipErrorKind::UnsupportedCompressionMethod(method) => {
            SourceContentPublicationError::Unsupported {
                reason: format!("unsupported generated compression method {method}"),
            }
        },
        ZipErrorKind::IO(source) | ZipErrorKind::Io(source) => {
            let default = if progress == SourceContentPublicationProgress::Untouched {
                UnmarkedIo::Archive
            } else {
                UnmarkedIo::Sink
            };
            map_owned_io_error(source, progress, default)
        },
        kind => SourceContentPublicationError::Archive {
            progress,
            source: soapberry_zip::Error::from(kind),
        },
    }
}

fn map_io_error(
    error: io::Error,
    progress: SourceContentPublicationProgress,
) -> SourceContentPublicationError {
    map_owned_io_error(error, progress, UnmarkedIo::Sink)
}

#[derive(Debug, Clone, Copy)]
enum UnmarkedIo {
    Source,
    Sink,
    Archive,
}

fn map_owned_io_error(
    error: io::Error,
    progress: SourceContentPublicationProgress,
    default: UnmarkedIo,
) -> SourceContentPublicationError {
    if let Some(failure) = error
        .get_ref()
        .and_then(|source| source.downcast_ref::<PublicationIoFailure>())
    {
        return map_publication_io_failure(failure, progress);
    }
    if let Some(changed) = error
        .get_ref()
        .and_then(|source| source.downcast_ref::<SourceChangedIo>())
    {
        return SourceContentPublicationError::SourceChanged {
            expected: changed.expected,
            observed: changed.observed,
            progress,
        };
    }
    if error
        .get_ref()
        .is_some_and(|source| source.is::<SourceExecutionIo>())
    {
        let kind = error.kind();
        if let Some(source) = error.into_inner() {
            match source.downcast::<SourceExecutionIo>() {
                Ok(source) => return map_execution_error(source.source, progress),
                Err(source) => {
                    return SourceContentPublicationError::Source {
                        progress,
                        source: Error::Io(io::Error::new(kind, source)),
                    };
                },
            }
        }
        return SourceContentPublicationError::Source {
            progress,
            source: Error::Io(io::Error::from(kind)),
        };
    }
    if error
        .get_ref()
        .is_some_and(|source| source.is::<SourceReadIo>())
    {
        let kind = error.kind();
        if let Some(source) = error.into_inner() {
            match source.downcast::<SourceReadIo>() {
                Ok(source) => {
                    return map_owned_io_error(source.source, progress, UnmarkedIo::Source);
                },
                Err(source) => {
                    return SourceContentPublicationError::Source {
                        progress,
                        source: Error::Io(io::Error::new(kind, source)),
                    };
                },
            }
        }
        return SourceContentPublicationError::Source {
            progress,
            source: Error::Io(io::Error::from(kind)),
        };
    }
    if error
        .get_ref()
        .is_some_and(|source| source.is::<SinkWriteIo>())
    {
        let kind = error.kind();
        if let Some(source) = error.into_inner() {
            match source.downcast::<SinkWriteIo>() {
                Ok(source) => {
                    return SourceContentPublicationError::Sink {
                        progress,
                        source: source.source,
                    };
                },
                Err(source) => {
                    return SourceContentPublicationError::Sink {
                        progress,
                        source: io::Error::new(kind, source),
                    };
                },
            }
        }
        return SourceContentPublicationError::Sink {
            progress,
            source: io::Error::from(kind),
        };
    }
    match default {
        UnmarkedIo::Source => SourceContentPublicationError::Source {
            progress,
            source: Error::Io(error),
        },
        UnmarkedIo::Sink => SourceContentPublicationError::Sink {
            progress,
            source: error,
        },
        UnmarkedIo::Archive => SourceContentPublicationError::Archive {
            progress,
            source: soapberry_zip::Error::from(error),
        },
    }
}

fn map_publication_io_failure(
    failure: &PublicationIoFailure,
    progress: SourceContentPublicationProgress,
) -> SourceContentPublicationError {
    match failure {
        PublicationIoFailure::Cancelled => SourceContentPublicationError::Cancelled { progress },
        PublicationIoFailure::Limit { actual, maximum } => {
            SourceContentPublicationError::LimitExceeded {
                progress,
                actual: *actual,
                maximum: *maximum,
            }
        },
        PublicationIoFailure::Execution(error) => map_execution_error(error.clone(), progress),
        PublicationIoFailure::SourceChanged { expected, observed } => {
            SourceContentPublicationError::SourceChanged {
                expected: *expected,
                observed: *observed,
                progress,
            }
        },
    }
}

fn publication_execution_failure(error: &ExecutionError) -> PublicationIoFailure {
    match error {
        ExecutionError::Cancelled => PublicationIoFailure::Cancelled,
        source => PublicationIoFailure::Execution(source.clone()),
    }
}

fn unsupported(reason: impl Into<String>) -> SourceContentPublicationError {
    SourceContentPublicationError::Unsupported {
        reason: reason.into(),
    }
}

fn le_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    bytes
        .get(offset..offset.checked_add(2)?)
        .and_then(|value| value.try_into().ok())
        .map(u16::from_le_bytes)
}

fn le_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    bytes
        .get(offset..offset.checked_add(4)?)
        .and_then(|value| value.try_into().ok())
        .map(u32::from_le_bytes)
}

fn le_u64(bytes: &[u8], offset: usize) -> Option<u64> {
    bytes
        .get(offset..offset.checked_add(8)?)
        .and_then(|value| value.try_into().ok())
        .map(u64::from_le_bytes)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::PackageWriter;
    use litchi_core::OwnedSource;
    use std::io::Cursor;

    const MIME: &str = constants::ODF_TEXT;
    const SOURCE: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body/></office:document-content>"#;
    const TARGET: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text/></office:body></office:document-content>"#;

    fn package_bytes() -> Vec<u8> {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(MIME).expect("mimetype");
        writer.add_file("content.xml", SOURCE).expect("content");
        writer
            .add_file_with_media_type("Pictures/blob.bin", b"opaque", "application/octet-stream")
            .expect("opaque");
        writer.finish_to_bytes().expect("package")
    }

    fn package_bytes_with_stale_manifest_size() -> Vec<u8> {
        let mut bytes = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut bytes);
            let stored = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            let deflated = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            writer
                .start_file("mimetype", stored)
                .expect("mimetype entry");
            writer.write_all(MIME.as_bytes()).expect("mimetype bytes");
            writer
                .start_file(constants::ODF_CONTENT, deflated)
                .expect("content entry");
            writer.write_all(SOURCE).expect("content bytes");
            writer
                .start_file(constants::ODF_MANIFEST, deflated)
                .expect("manifest entry");
            writer
                .write_all(
                    format!(
                        r#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="{}" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="{}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml" manifest:size="1"/></manifest:manifest>"#,
                        "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
                        MIME,
                    )
                    .as_bytes(),
                )
                .expect("manifest bytes");
            writer.finish().expect("finish ZIP");
        }
        bytes.into_inner()
    }

    #[test]
    fn changed_publication_reports_output_and_raw_reopens() {
        let bytes = package_bytes();
        let source = Arc::new(OwnedSource::new(bytes.clone()));
        let package = SourceBackedPackage::from_read_at(source).expect("open");
        let mut output = Vec::new();
        let report = package
            .write_content_xml_to_stream(&mut output, TARGET)
            .expect("publish");
        assert!(!report.is_no_op());
        assert_eq!(report.bytes(), output.len() as u64);
        assert_ne!(output, bytes);
    }

    #[test]
    fn exact_no_op_copies_source_identity() {
        let bytes = package_bytes();
        let source = Arc::new(OwnedSource::new(bytes.clone()));
        let package = SourceBackedPackage::from_read_at(source).expect("open");
        let mut output = Vec::new();
        let report = package
            .write_content_xml_to_stream(&mut output, SOURCE)
            .expect("publish");
        assert!(report.is_no_op());
        assert_eq!(output, bytes);
    }

    #[test]
    fn known_content_accepts_stale_unencrypted_manifest_size_only_for_exact_noop() {
        let bytes = package_bytes_with_stale_manifest_size();
        let package = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes.clone())))
            .expect("open source with stale manifest:size");
        let (known, proof) = package
            .get_content_xml_with_source_proof()
            .expect("authenticate ZIP-declared content size");
        assert_eq!(known, SOURCE);

        let mut output = Vec::new();
        let report = package
            .write_content_xml_to_stream_with_known_source_content(
                &mut output,
                &known,
                &proof,
                SOURCE,
                SourceContentPublicationOptions::default(),
            )
            .expect("exact no-op preserves stale producer metadata");
        assert!(report.is_no_op());
        assert_eq!(output, bytes);

        let mut output = Vec::new();
        let error = package
            .write_content_xml_to_stream_with_known_source_content(
                &mut output,
                &known,
                &proof,
                TARGET,
                SourceContentPublicationOptions::default(),
            )
            .expect_err("changed publication must refuse manifest:size");
        assert!(matches!(
            error,
            SourceContentPublicationError::Unsupported { reason }
                if reason.contains("manifest:size")
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn known_source_content_rejects_mismatched_bytes_and_size_before_output() {
        let bytes = package_bytes();
        let package =
            SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).expect("open");
        let (known, proof) = package
            .get_content_xml_with_source_proof()
            .expect("authenticated content");

        let mut mismatched = known.clone();
        mismatched[0] ^= 1;
        let mut output = Vec::new();
        let error = package
            .write_content_xml_to_stream_with_known_source_content(
                &mut output,
                &mismatched,
                &proof,
                TARGET,
                SourceContentPublicationOptions::default(),
            )
            .expect_err("mismatched retained bytes must fail closed");
        assert!(matches!(
            error,
            SourceContentPublicationError::Core(Error::InvalidFormat(message))
                if message.contains("authenticated source member")
        ));
        assert!(output.is_empty());

        let mut output = Vec::new();
        let error = package
            .write_content_xml_to_stream_with_known_source_content(
                &mut output,
                &known[..known.len() - 1],
                &proof,
                TARGET,
                SourceContentPublicationOptions::default(),
            )
            .expect_err("mismatched retained size must fail closed");
        assert!(matches!(
            error,
            SourceContentPublicationError::Core(Error::InvalidFormat(message))
                if message.contains("indexed ZIP size")
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn known_source_content_exact_no_op_preserves_signed_like_archive_bytes() {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(MIME).expect("mimetype");
        writer.add_file("content.xml", SOURCE).expect("content");
        writer
            .add_file_with_media_type(
                "META-INF/documentsignatures.xml",
                br#"<signatures xmlns="urn:test"/>"#,
                "application/vnd.oasis.opendocument.digital-signature",
            )
            .expect("signature");
        let bytes = writer.finish_to_bytes().expect("package");
        let package =
            SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
        let (known, proof) = package
            .get_content_xml_with_source_proof()
            .expect("authenticated content");
        let mut output = Vec::new();
        let report = package
            .write_content_xml_to_stream_with_known_source_content(
                &mut output,
                &known,
                &proof,
                SOURCE,
                SourceContentPublicationOptions::default(),
            )
            .expect("exact no-op");
        assert!(report.is_no_op());
        assert_eq!(output, bytes);
    }

    #[test]
    fn known_source_content_refuses_encrypted_source_before_output() {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(MIME).expect("mimetype");
        writer
            .set_encryption("source-password", crate::core::Profile::compatible())
            .expect("encryption");
        writer.add_file("content.xml", SOURCE).expect("content");
        let bytes = writer.finish_to_bytes().expect("package");
        let package = SourceBackedPackage::from_read_at_with_password(
            Arc::new(OwnedSource::new(bytes)),
            "source-password",
        )
        .expect("open");
        let (known, proof) = package
            .get_content_xml_with_source_proof()
            .expect("authenticated decrypted content");
        let mut output = Vec::new();
        let error = package
            .write_content_xml_to_stream_with_known_source_content(
                &mut output,
                &known,
                &proof,
                TARGET,
                SourceContentPublicationOptions::default(),
            )
            .expect_err("encrypted source must remain refused");
        assert!(matches!(
            error,
            SourceContentPublicationError::Unsupported { .. }
        ));
        assert!(output.is_empty());
    }
}
