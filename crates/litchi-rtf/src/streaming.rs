//! Bounded, forward-only authoring of small plain-text RTF documents.
//!
//! [`StreamingRtfWriter`] is deliberately narrower than [`crate::write::Writer`].
//! It creates a fresh document whose body is made up of explicitly ordered
//! paragraphs and plain runs.  The caller owns the sequencing: no seekable
//! output, document-wide model, or deferred final serialization is required.
//! A run implements [`std::io::Write`] and accepts UTF-8 bytes in chunks.

#![allow(
    clippy::indexing_slicing,
    reason = "fixed-size escape buffers are populated only after their bounded lengths are established"
)]

use std::fmt;
use std::io::{self, Write};

use litchi_codepage::Mbcs;
use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource};

use crate::write::Charset;

// Four bytes hold one incomplete UTF-8 scalar, 32 bytes hold the largest
// escaped scalar, and one byte records a pending CR for CRLF normalization.
const FIXED_SCRATCH_BYTES: u64 = 37;
// ASCII body writes are bounded to the fixed escape-buffer scale. Cancellation
// is checked between these chunks, so one full sink call can accept at most 32
// body bytes before cancellation is observed.
const MAX_ASCII_BATCH_BYTES: usize = 32;
const DEFAULT_MAX_BYTES: u64 = 256 * 1024 * 1024;
const DEFAULT_MAX_PARAGRAPHS: u64 = 1_000_000;
const DEFAULT_MAX_RUNS: u64 = 4_000_000;
const HEX: &[u8; 16] = b"0123456789abcdef";

/// Finite limits for [`StreamingRtfWriter`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StreamingRtfLimits {
    /// Maximum UTF-8 input bytes accepted by all runs.
    pub max_input_bytes: u64,
    /// Maximum bytes accepted by the output sink.
    pub max_output_bytes: u64,
    /// Maximum number of paragraphs.
    pub max_paragraphs: u64,
    /// Maximum number of runs.
    pub max_runs: u64,
    /// Bounded scratch reservation retained for the writer's UTF-8/control
    /// encoding path.
    pub max_scratch_bytes: u64,
}

impl StreamingRtfLimits {
    /// Creates an explicit finite limit set.
    #[must_use]
    pub const fn new(
        max_input_bytes: u64,
        max_output_bytes: u64,
        max_paragraphs: u64,
        max_runs: u64,
        max_scratch_bytes: u64,
    ) -> Self {
        Self {
            max_input_bytes,
            max_output_bytes,
            max_paragraphs,
            max_runs,
            max_scratch_bytes,
        }
    }

    fn validate(self) -> Result<(), StreamingRtfError> {
        if self.max_scratch_bytes < FIXED_SCRATCH_BYTES {
            return Err(StreamingRtfError::InvalidInput {
                message: format!(
                    "streaming RTF scratch limit must be at least {FIXED_SCRATCH_BYTES} bytes"
                ),
                written: 0,
            });
        }
        usize::try_from(self.max_scratch_bytes).map_err(|_error| {
            StreamingRtfError::InvalidInput {
                message: "streaming RTF scratch limit exceeds usize".to_string(),
                written: 0,
            }
        })?;
        Ok(())
    }
}

impl Default for StreamingRtfLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: DEFAULT_MAX_BYTES,
            max_output_bytes: DEFAULT_MAX_BYTES,
            max_paragraphs: DEFAULT_MAX_PARAGRAPHS,
            max_runs: DEFAULT_MAX_RUNS,
            max_scratch_bytes: 64,
        }
    }
}

/// Header choices for [`StreamingRtfWriter`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StreamingRtfOptions {
    /// Legacy code page declaration and escape policy.
    pub charset: Charset,
    /// Default font index declared by `\\deff`.
    pub default_font: u16,
}

impl Default for StreamingRtfOptions {
    fn default() -> Self {
        Self {
            charset: Charset::WINDOWS_1252,
            default_font: 0,
        }
    }
}

/// Failure from a bounded forward-only RTF authoring operation.
#[derive(Debug)]
#[non_exhaustive]
pub enum StreamingRtfError {
    /// The sink failed after accepting the reported number of output bytes.
    Io {
        /// Underlying I/O category.
        kind: io::ErrorKind,
        /// Stable diagnostic captured from the sink.
        message: String,
        /// Output bytes accepted before the failure.
        written: u64,
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
}

impl StreamingRtfError {
    /// Output bytes accepted before this failure.
    #[must_use]
    pub const fn written(&self) -> u64 {
        match self {
            Self::Io { written, .. }
            | Self::InvalidInput { written, .. }
            | Self::LimitExceeded { written, .. }
            | Self::Cancelled { written } => *written,
        }
    }
}

impl fmt::Display for StreamingRtfError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Io {
                kind,
                message,
                written,
            } => write!(
                formatter,
                "streaming RTF sink failed ({kind}) after {written} output bytes: {message}"
            ),
            Self::InvalidInput { message, written } => write!(
                formatter,
                "invalid streaming RTF authoring input after {written} output bytes: {message}"
            ),
            Self::LimitExceeded {
                resource,
                observed,
                limit,
                written,
            } => write!(
                formatter,
                "streaming RTF limit exceeded for {resource}: observed {observed}, limit {limit}, output {written}"
            ),
            Self::Cancelled { written } => write!(
                formatter,
                "streaming RTF authoring cancelled after {written} output bytes"
            ),
        }
    }
}

impl std::error::Error for StreamingRtfError {}

#[derive(Debug, Clone)]
enum PoisonReason {
    Io {
        kind: io::ErrorKind,
        message: String,
    },
    InvalidInput(String),
    Limit {
        resource: &'static str,
        observed: u64,
        limit: u64,
    },
    Cancelled,
}

impl PoisonReason {
    fn error(&self, written: u64) -> StreamingRtfError {
        match self {
            Self::Io { kind, message } => StreamingRtfError::Io {
                kind: *kind,
                message: message.clone(),
                written,
            },
            Self::InvalidInput(message) => StreamingRtfError::InvalidInput {
                message: message.clone(),
                written,
            },
            Self::Limit {
                resource,
                observed,
                limit,
            } => StreamingRtfError::LimitExceeded {
                resource,
                observed: *observed,
                limit: *limit,
                written,
            },
            Self::Cancelled => StreamingRtfError::Cancelled { written },
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Phase {
    Ready,
    Paragraph { run_open: bool },
}

/// Sequential authoring state for a fresh plain-text RTF document.
pub struct StreamingRtfWriter<W: Write> {
    writer: W,
    context: ExecutionContext,
    limits: StreamingRtfLimits,
    options: StreamingRtfOptions,
    phase: Phase,
    paragraphs: u64,
    runs: u64,
    input_bytes: u64,
    output_bytes: u64,
    pending_utf8: [u8; 4],
    pending_len: usize,
    pending_cr: bool,
    _scratch: Reservation,
    poison: Option<PoisonReason>,
}

impl<W: Write> StreamingRtfWriter<W> {
    /// Creates a writer and immediately publishes its deterministic header.
    ///
    /// The sink only needs [`Write`].  The returned writer is ready for
    /// [`Self::start_paragraph`], and an empty document can be completed with
    /// [`Self::finish`].
    pub fn new(
        writer: W,
        context: ExecutionContext,
        limits: StreamingRtfLimits,
    ) -> Result<Self, StreamingRtfError> {
        Self::with_options(writer, context, limits, StreamingRtfOptions::default())
    }

    /// Creates a writer with an explicit legacy charset and default font.
    pub fn with_options(
        writer: W,
        context: ExecutionContext,
        limits: StreamingRtfLimits,
        options: StreamingRtfOptions,
    ) -> Result<Self, StreamingRtfError> {
        limits.validate()?;
        context.check().map_err(|error| map_execution(error, 0))?;
        let scratch = context
            .reserve(Resource::Memory, FIXED_SCRATCH_BYTES)
            .map_err(|error| map_execution(error, 0))?;
        let mut result = Self {
            writer,
            context,
            limits,
            options,
            phase: Phase::Ready,
            paragraphs: 0,
            runs: 0,
            input_bytes: 0,
            output_bytes: 0,
            pending_utf8: [0; 4],
            pending_len: 0,
            pending_cr: false,
            _scratch: scratch,
            poison: None,
        };
        emit_header(&mut result)?;
        Ok(result)
    }

    /// Number of paragraphs that have been started.
    #[must_use]
    pub const fn paragraph_count(&self) -> u64 {
        self.paragraphs
    }

    /// Number of runs that have been started.
    #[must_use]
    pub const fn run_count(&self) -> u64 {
        self.runs
    }

    /// UTF-8 input bytes accepted by the run writer.
    #[must_use]
    pub const fn input_bytes(&self) -> u64 {
        self.input_bytes
    }

    /// Output bytes accepted by the caller's sink.
    #[must_use]
    pub const fn output_bytes(&self) -> u64 {
        self.output_bytes
    }

    /// Whether an unrecoverable sink, cancellation, or resource failure has
    /// invalidated this writer.
    #[must_use]
    pub const fn is_poisoned(&self) -> bool {
        self.poison.is_some()
    }

    /// Start the next logical paragraph.
    pub fn start_paragraph(&mut self) -> Result<(), StreamingRtfError> {
        self.ensure_usable()?;
        if !matches!(self.phase, Phase::Ready) {
            return Err(self.state_error("a paragraph is already open"));
        }
        self.context
            .check()
            .map_err(|error| self.fail_execution(error))?;
        let observed = self.paragraphs.saturating_add(1);
        if observed > self.limits.max_paragraphs {
            return Err(self.fail_limit("paragraphs", observed, self.limits.max_paragraphs));
        }
        self.context
            .consume(Resource::Objects, 1)
            .and_then(|_| self.context.consume(Resource::Work, 1))
            .map_err(|error| self.fail_execution(error))?;
        self.emit(b"\\pard ")?;
        self.paragraphs = observed;
        self.phase = Phase::Paragraph { run_open: false };
        Ok(())
    }

    /// Start a plain run in the currently open paragraph.
    pub fn start_run(&mut self) -> Result<(), StreamingRtfError> {
        self.ensure_usable()?;
        if !matches!(self.phase, Phase::Paragraph { run_open: false }) {
            return Err(self.state_error("a paragraph must be open without another run"));
        }
        self.context
            .check()
            .map_err(|error| self.fail_execution(error))?;
        let observed = self.runs.saturating_add(1);
        if observed > self.limits.max_runs {
            return Err(self.fail_limit("runs", observed, self.limits.max_runs));
        }
        self.context
            .consume(Resource::Objects, 1)
            .and_then(|_| self.context.consume(Resource::Work, 1))
            .map_err(|error| self.fail_execution(error))?;
        self.emit(b"{\\plain ")?;
        self.runs = observed;
        self.phase = Phase::Paragraph { run_open: true };
        Ok(())
    }

    /// Finish the current run.
    pub fn finish_run(&mut self) -> Result<(), StreamingRtfError> {
        self.ensure_usable()?;
        if !matches!(self.phase, Phase::Paragraph { run_open: true }) {
            return Err(self.state_error("no run is open"));
        }
        if self.pending_len != 0 {
            return Err(self.fail_invalid("the current run ends with incomplete UTF-8"));
        }
        self.flush_pending_cr()?;
        self.emit(b"}")?;
        self.phase = Phase::Paragraph { run_open: false };
        Ok(())
    }

    /// Finish the current paragraph.
    pub fn finish_paragraph(&mut self) -> Result<(), StreamingRtfError> {
        self.ensure_usable()?;
        if !matches!(self.phase, Phase::Paragraph { run_open: false }) {
            return Err(self.state_error("finish the open run before the paragraph"));
        }
        if self.pending_len != 0 {
            return Err(self.fail_invalid("the current paragraph ends with incomplete UTF-8"));
        }
        self.flush_pending_cr()?;
        self.emit(b"\\par ")?;
        self.phase = Phase::Ready;
        Ok(())
    }

    /// Finish the document and return the owned non-seek sink.
    pub fn finish(mut self) -> Result<W, StreamingRtfError> {
        self.ensure_usable()?;
        if !matches!(self.phase, Phase::Ready) {
            return Err(self.state_error("finish the current paragraph first"));
        }
        if self.pending_len != 0 {
            return Err(self.fail_invalid("the document ends with incomplete UTF-8"));
        }
        self.flush_pending_cr()?;
        self.emit(b"}")?;
        Ok(self.writer)
    }

    fn ensure_usable(&self) -> Result<(), StreamingRtfError> {
        self.poison
            .as_ref()
            .map_or(Ok(()), |reason| Err(reason.error(self.output_bytes)))
    }

    fn state_error(&self, message: &str) -> StreamingRtfError {
        StreamingRtfError::InvalidInput {
            message: message.to_string(),
            written: self.output_bytes,
        }
    }

    fn fail_invalid(&mut self, message: &str) -> StreamingRtfError {
        self.poison = Some(PoisonReason::InvalidInput(message.to_string()));
        self.poison.as_ref().map_or_else(
            || self.state_error(message),
            |reason| reason.error(self.output_bytes),
        )
    }

    fn fail_limit(
        &mut self,
        resource: &'static str,
        observed: u64,
        limit: u64,
    ) -> StreamingRtfError {
        self.poison = Some(PoisonReason::Limit {
            resource,
            observed,
            limit,
        });
        self.poison.as_ref().map_or_else(
            || self.state_error("streaming RTF limit exceeded"),
            |reason| reason.error(self.output_bytes),
        )
    }

    fn fail_execution(&mut self, error: ExecutionError) -> StreamingRtfError {
        let reason = poison_from_execution(error);
        self.poison = Some(reason);
        self.poison.as_ref().map_or_else(
            || self.state_error("streaming RTF execution failure"),
            |reason| reason.error(self.output_bytes),
        )
    }

    fn fail_io(&mut self, error: io::Error) -> StreamingRtfError {
        self.poison = Some(PoisonReason::Io {
            kind: error.kind(),
            message: error.to_string(),
        });
        self.poison.as_ref().map_or_else(
            || self.state_error("streaming RTF sink failure"),
            |reason| reason.error(self.output_bytes),
        )
    }

    fn commit_output(&mut self, reservation: Reservation, written: usize) {
        let accepted = u64::try_from(written).unwrap_or(u64::MAX);
        let _ = reservation.commit(accepted);
        self.output_bytes = self.output_bytes.saturating_add(accepted);
    }

    fn emit(&mut self, bytes: &[u8]) -> Result<(), StreamingRtfError> {
        self.ensure_usable()?;
        if bytes.is_empty() {
            return Ok(());
        }
        let requested = u64::try_from(bytes.len()).unwrap_or(u64::MAX);
        let observed = self.output_bytes.saturating_add(requested);
        if observed > self.limits.max_output_bytes {
            return Err(self.fail_limit("output bytes", observed, self.limits.max_output_bytes));
        }
        self.context
            .check()
            .map_err(|error| self.fail_execution(error))?;
        let reservation = self
            .context
            .reserve(Resource::OutputBytes, requested)
            .map_err(|error| self.fail_execution(error))?;
        let mut written = 0usize;
        loop {
            if let Err(error) = self.context.check() {
                self.commit_output(reservation, written);
                return Err(self.fail_execution(error));
            }
            match self.writer.write(&bytes[written..]) {
                Ok(0) => {
                    self.commit_output(reservation, written);
                    return Err(self.fail_io(io::Error::new(
                        io::ErrorKind::WriteZero,
                        "streaming RTF sink returned zero progress",
                    )));
                },
                Ok(count) => {
                    if count > bytes.len().saturating_sub(written) {
                        self.commit_output(reservation, written);
                        return Err(
                            self.fail_invalid("streaming RTF sink reported excess progress")
                        );
                    }
                    written = written.saturating_add(count);
                    if written >= bytes.len() {
                        self.commit_output(reservation, written);
                        return Ok(());
                    }
                },
                Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
                Err(error) => {
                    self.commit_output(reservation, written);
                    return Err(self.fail_io(error));
                },
            }
        }
    }

    fn emit_decimal(&mut self, value: u32) -> Result<(), StreamingRtfError> {
        let mut divisor = 1_u32;
        while value / divisor >= 10 {
            divisor = divisor.saturating_mul(10);
        }
        while divisor != 0 {
            let digit = b'0' + u8::try_from((value / divisor) % 10).unwrap_or(0);
            self.emit(&[digit])?;
            divisor /= 10;
        }
        Ok(())
    }

    fn process_valid_text(
        &mut self,
        text: &str,
        accepted: &mut usize,
        input_already_counted: bool,
    ) -> Result<(), StreamingRtfError> {
        let bytes = text.as_bytes();
        let mut offset = 0usize;
        while offset < bytes.len() {
            if !self.pending_cr && is_batchable_ascii(bytes[offset]) {
                let start = offset;
                offset = offset.saturating_add(1);
                while offset < bytes.len() && is_batchable_ascii(bytes[offset]) {
                    offset = offset.saturating_add(1);
                }
                let mut chunk_start = start;
                while chunk_start < offset {
                    let chunk_end = chunk_start
                        .saturating_add(MAX_ASCII_BATCH_BYTES)
                        .min(offset);
                    let span = bytes
                        .get(chunk_start..chunk_end)
                        .ok_or_else(|| self.fail_invalid("ASCII span is outside the input"))?;
                    if !self.try_emit_ascii_span(span, accepted, input_already_counted)? {
                        for byte in span {
                            if !input_already_counted {
                                *accepted = accepted.saturating_add(1);
                            }
                            self.emit_character(char::from(*byte))?;
                        }
                    }
                    chunk_start = chunk_end;
                }
                continue;
            }

            let character = text
                .get(offset..)
                .and_then(|remaining| remaining.chars().next())
                .ok_or_else(|| self.fail_invalid("UTF-8 character is outside the input"))?;
            offset = offset.saturating_add(character.len_utf8());
            if !input_already_counted {
                *accepted = accepted.saturating_add(character.len_utf8());
            }
            self.emit_character(character)?;
        }
        Ok(())
    }

    fn try_emit_ascii_span(
        &mut self,
        span: &[u8],
        accepted: &mut usize,
        input_already_counted: bool,
    ) -> Result<bool, StreamingRtfError> {
        if span.len() < 2 || span.len() > MAX_ASCII_BATCH_BYTES {
            return Ok(false);
        }
        let requested = u64::try_from(span.len()).unwrap_or(u64::MAX);
        let observed = self.output_bytes.saturating_add(requested);
        if observed > self.limits.max_output_bytes {
            return Ok(false);
        }

        // Work is cumulative while output is outstanding. Holding both
        // reservations makes the span atomic against local and hierarchical
        // limits; failure of either reservation falls back to scalar output,
        // which retains the original per-character boundary and precedence.
        let work = match self.context.reserve(Resource::Work, requested) {
            Ok(reservation) => reservation,
            Err(_error) => return Ok(false),
        };
        let output = match self.context.reserve(Resource::OutputBytes, requested) {
            Ok(reservation) => reservation,
            Err(_error) => {
                drop(work);
                return Ok(false);
            },
        };
        self.emit_reserved_ascii(span, accepted, input_already_counted, work, output)?;
        Ok(true)
    }

    fn emit_reserved_ascii(
        &mut self,
        bytes: &[u8],
        accepted: &mut usize,
        input_already_counted: bool,
        work: Reservation,
        output: Reservation,
    ) -> Result<(), StreamingRtfError> {
        let requested = bytes.len();
        let requested_u64 = u64::try_from(requested).unwrap_or(u64::MAX);
        let mut written = 0usize;
        loop {
            if let Err(error) = self.context.check() {
                self.commit_reserved_ascii(
                    work,
                    output,
                    written,
                    requested,
                    accepted,
                    input_already_counted,
                );
                return Err(self.fail_execution(error));
            }
            match self.writer.write(&bytes[written..]) {
                Ok(0) => {
                    self.commit_reserved_ascii(
                        work,
                        output,
                        written,
                        requested,
                        accepted,
                        input_already_counted,
                    );
                    return Err(self.fail_io(io::Error::new(
                        io::ErrorKind::WriteZero,
                        "streaming RTF sink returned zero progress",
                    )));
                },
                Ok(count) => {
                    if count > bytes.len().saturating_sub(written) {
                        self.commit_reserved_ascii(
                            work,
                            output,
                            written,
                            requested,
                            accepted,
                            input_already_counted,
                        );
                        return Err(
                            self.fail_invalid("streaming RTF sink reported excess progress")
                        );
                    }
                    written = written.saturating_add(count);
                    if let Err(error) = self.context.check() {
                        self.commit_reserved_ascii(
                            work,
                            output,
                            written,
                            requested,
                            accepted,
                            input_already_counted,
                        );
                        return Err(self.fail_execution(error));
                    }
                    if written >= bytes.len() {
                        let _ = work.commit(requested_u64);
                        let _ = output.commit(requested_u64);
                        if !input_already_counted {
                            *accepted = accepted.saturating_add(requested);
                        }
                        self.output_bytes = self.output_bytes.saturating_add(requested_u64);
                        return Ok(());
                    }
                },
                Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
                Err(error) => {
                    self.commit_reserved_ascii(
                        work,
                        output,
                        written,
                        requested,
                        accepted,
                        input_already_counted,
                    );
                    return Err(self.fail_io(error));
                },
            }
        }
    }

    fn commit_reserved_ascii(
        &mut self,
        work: Reservation,
        output: Reservation,
        written: usize,
        requested: usize,
        accepted: &mut usize,
        input_already_counted: bool,
    ) {
        let attempted = written.saturating_add(1).min(requested);
        let _ = work.commit(u64::try_from(attempted).unwrap_or(u64::MAX));
        let _ = output.commit(u64::try_from(written).unwrap_or(u64::MAX));
        if !input_already_counted {
            *accepted = accepted.saturating_add(attempted);
        }
        self.output_bytes = self
            .output_bytes
            .saturating_add(u64::try_from(written).unwrap_or(u64::MAX));
    }

    fn write_utf8(&mut self, bytes: &[u8]) -> Result<usize, StreamingRtfError> {
        if bytes.is_empty() {
            return Ok(0);
        }
        if !matches!(self.phase, Phase::Paragraph { run_open: true }) {
            return Err(self.state_error("write bytes only while a run is open"));
        }
        self.ensure_usable()?;
        let observed = self
            .input_bytes
            .saturating_add(u64::try_from(bytes.len()).unwrap_or(u64::MAX));
        if observed > self.limits.max_input_bytes {
            return Err(self.fail_limit("input bytes", observed, self.limits.max_input_bytes));
        }
        let requested = u64::try_from(bytes.len()).unwrap_or(u64::MAX);
        let reservation = self
            .context
            .reserve(Resource::InputBytes, requested)
            .map_err(|error| self.fail_execution(error))?;
        let mut accepted = 0usize;
        let result = self.process_utf8(bytes, &mut accepted);
        let _ = reservation.commit(u64::try_from(accepted).unwrap_or(u64::MAX));
        self.input_bytes = self
            .input_bytes
            .saturating_add(u64::try_from(accepted).unwrap_or(u64::MAX));
        result.map(|()| accepted)
    }

    fn process_utf8(
        &mut self,
        bytes: &[u8],
        accepted: &mut usize,
    ) -> Result<(), StreamingRtfError> {
        let mut offset = 0usize;
        if self.pending_len != 0 {
            let expected = utf8_sequence_len(self.pending_utf8[0])
                .ok_or_else(|| self.fail_invalid("input contains malformed UTF-8"))?;
            let needed = expected.saturating_sub(self.pending_len);
            let count = needed.min(bytes.len());
            let end = self.pending_len.saturating_add(count);
            self.pending_utf8[self.pending_len..end].copy_from_slice(&bytes[..count]);
            self.pending_len = end;
            offset = count;
            *accepted = accepted.saturating_add(count);
            if self.pending_len < expected {
                return Ok(());
            }
            let completed = self.pending_utf8;
            match std::str::from_utf8(&completed[..expected]) {
                Ok(text) => {
                    self.pending_len = 0;
                    self.process_valid_text(text, accepted, true)?;
                },
                Err(_error) => {
                    self.pending_len = 0;
                    return Err(self.fail_invalid("input contains malformed UTF-8"));
                },
            }
        }

        if offset >= bytes.len() {
            return Ok(());
        }
        let remaining = bytes.get(offset..).ok_or_else(|| {
            self.fail_invalid("UTF-8 input offset is outside the provided buffer")
        })?;
        match std::str::from_utf8(remaining) {
            Ok(text) => self.process_valid_text(text, accepted, false),
            Err(error) => {
                let valid = error.valid_up_to();
                if let Some(text_bytes) = remaining.get(..valid) {
                    let text = std::str::from_utf8(text_bytes).map_err(|_error| {
                        self.fail_invalid("UTF-8 prefix is unexpectedly malformed")
                    })?;
                    self.process_valid_text(text, accepted, false)?;
                }
                if error.error_len().is_none() {
                    let tail = remaining.get(valid..).ok_or_else(|| {
                        self.fail_invalid("UTF-8 tail is outside the provided buffer")
                    })?;
                    if tail.len() > self.pending_utf8.len() {
                        return Err(self.fail_invalid("UTF-8 sequence exceeds four bytes"));
                    }
                    self.pending_utf8[..tail.len()].copy_from_slice(tail);
                    self.pending_len = tail.len();
                    *accepted = accepted.saturating_add(tail.len());
                    return Ok(());
                }
                Err(self.fail_invalid("input contains malformed UTF-8"))
            },
        }
    }

    fn emit_character(&mut self, character: char) -> Result<(), StreamingRtfError> {
        self.context
            .consume(Resource::Work, 1)
            .map_err(|error| self.fail_execution(error))?;
        if self.pending_cr {
            if character == '\n' {
                self.pending_cr = false;
                self.emit(b"\\line ")?;
                return Ok(());
            }
            self.pending_cr = false;
            self.emit(b"\\line ")?;
        }
        if character == '\r' {
            self.pending_cr = true;
            return Ok(());
        }
        let mut encoded = [0u8; 32];
        let length = encode_character(character, self.options.charset, &mut encoded);
        self.emit(&encoded[..length])
    }

    fn flush_pending_cr(&mut self) -> Result<(), StreamingRtfError> {
        if self.pending_cr {
            self.pending_cr = false;
            self.emit(b"\\line ")?;
        }
        Ok(())
    }
}

impl<W: Write> Write for StreamingRtfWriter<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        match self.write_utf8(bytes) {
            Ok(count) => Ok(count),
            Err(error) => Err(io::Error::new(io_kind(&error), error.to_string())),
        }
    }

    fn flush(&mut self) -> io::Result<()> {
        self.ensure_usable()
            .map_err(|error| io::Error::new(io_kind(&error), error.to_string()))?;
        if let Err(error) = self.context.check() {
            let mapped = self.fail_execution(error);
            return Err(io::Error::new(io_kind(&mapped), mapped.to_string()));
        }
        self.writer.flush().map_err(|error| {
            let mapped = self.fail_io(error);
            io::Error::new(io_kind(&mapped), mapped.to_string())
        })
    }
}

fn io_kind(error: &StreamingRtfError) -> io::ErrorKind {
    match error {
        StreamingRtfError::Io { kind, .. } => *kind,
        StreamingRtfError::InvalidInput { .. } => io::ErrorKind::InvalidInput,
        StreamingRtfError::LimitExceeded { .. } => io::ErrorKind::QuotaExceeded,
        StreamingRtfError::Cancelled { .. } => io::ErrorKind::Interrupted,
    }
}

fn map_execution(error: ExecutionError, written: u64) -> StreamingRtfError {
    match poison_from_execution(error) {
        PoisonReason::Limit {
            resource,
            observed,
            limit,
        } => StreamingRtfError::LimitExceeded {
            resource,
            observed,
            limit,
            written,
        },
        PoisonReason::Cancelled => StreamingRtfError::Cancelled { written },
        PoisonReason::Io { kind, message } => StreamingRtfError::Io {
            kind,
            message,
            written,
        },
        PoisonReason::InvalidInput(message) => StreamingRtfError::InvalidInput { message, written },
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

const fn resource_name(resource: Resource) -> &'static str {
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

fn is_batchable_ascii(byte: u8) -> bool {
    matches!(byte, 0x20..=0x7e) && !matches!(byte, b'\\' | b'{' | b'}')
}

fn encode_character(character: char, charset: Charset, output: &mut [u8; 32]) -> usize {
    match character {
        '\\' => copy_bytes(output, b"\\\\"),
        '{' => copy_bytes(output, b"\\{"),
        '}' => copy_bytes(output, b"\\}"),
        '\n' => copy_bytes(output, b"\\line "),
        '\r' => copy_bytes(output, b"\\line "),
        '\t' => copy_bytes(output, b"\\tab "),
        c if c.is_ascii() && (c as u32) >= 0x20 && (c as u32) <= 0x7e => copy_char(output, c),
        c if c.is_ascii() => {
            let mut length = 0usize;
            output[length] = b'\\';
            length += 1;
            output[length] = b'\'';
            length += 1;
            output[length] = HEX[((c as u8) >> 4) as usize];
            length += 1;
            output[length] = HEX[(c as u8 & 0x0f) as usize];
            length + 1
        },
        c => encode_non_ascii(c, charset, output),
    }
}

fn utf8_sequence_len(first: u8) -> Option<usize> {
    match first {
        0x00..=0x7f => Some(1),
        0xc2..=0xdf => Some(2),
        0xe0..=0xef => Some(3),
        0xf0..=0xf4 => Some(4),
        _ => None,
    }
}

fn encode_non_ascii(character: char, charset: Charset, output: &mut [u8; 32]) -> usize {
    if let Some(byte) = cp1252_byte(character, charset) {
        let mut length = 0usize;
        output[length] = b'\\';
        length += 1;
        output[length] = b'\'';
        length += 1;
        output[length] = HEX[(byte >> 4) as usize];
        length += 1;
        output[length] = HEX[(byte & 0x0f) as usize];
        length += 1;
        return length;
    }
    let mut length = 0usize;
    let mut units = [0u16; 2];
    let unit_count = character.encode_utf16(&mut units).len();
    for unit in units[..unit_count].iter().copied() {
        output[length] = b'\\';
        length += 1;
        output[length] = b'u';
        length += 1;
        let signed = i32::from(i16::from_ne_bytes(unit.to_ne_bytes()));
        length += write_decimal(&mut output[length..], signed);
    }
    for _unit in units[..unit_count].iter() {
        output[length] = b'?';
        length += 1;
    }
    length
}

fn cp1252_byte(character: char, charset: Charset) -> Option<u8> {
    if !matches!(charset, Charset::Ansi(page) if page == Mbcs::WINDOWS_1252) {
        return None;
    }
    match character as u32 {
        0xa0..=0xff => Some(character as u8),
        0x20ac => Some(0x80),
        0x201a => Some(0x82),
        0x192 => Some(0x83),
        0x201e => Some(0x84),
        0x2026 => Some(0x85),
        0x2020 => Some(0x86),
        0x2021 => Some(0x87),
        0x2c6 => Some(0x88),
        0x2030 => Some(0x89),
        0x160 => Some(0x8a),
        0x2039 => Some(0x8b),
        0x152 => Some(0x8c),
        0x17d => Some(0x8e),
        0x2018 => Some(0x91),
        0x2019 => Some(0x92),
        0x201c => Some(0x93),
        0x201d => Some(0x94),
        0x2022 => Some(0x95),
        0x2013 => Some(0x96),
        0x2014 => Some(0x97),
        0x2dc => Some(0x98),
        0x2122 => Some(0x99),
        0x161 => Some(0x9a),
        0x203a => Some(0x9b),
        0x153 => Some(0x9c),
        0x17e => Some(0x9e),
        0x178 => Some(0x9f),
        _ => None,
    }
}

fn write_decimal(output: &mut [u8], value: i32) -> usize {
    let negative = value < 0;
    let mut magnitude = if negative {
        i64::from(value).saturating_neg()
    } else {
        i64::from(value)
    };
    let mut offset = 0usize;
    if negative {
        output[offset] = b'-';
        offset += 1;
    }
    let start = offset;
    loop {
        output[offset] = b'0' + u8::try_from(magnitude % 10).unwrap_or(0);
        offset += 1;
        magnitude /= 10;
        if magnitude == 0 {
            break;
        }
    }
    output[start..offset].reverse();
    offset
}

fn emit_header<W: Write>(writer: &mut StreamingRtfWriter<W>) -> Result<(), StreamingRtfError> {
    writer.emit(b"{\\rtf1")?;
    match writer.options.charset {
        Charset::Ansi(page) => {
            writer.emit(b"\\ansi\\ansicpg")?;
            writer.emit_decimal(page.id())?;
        },
        Charset::Mac => writer.emit(b"\\mac")?,
        Charset::Pc => writer.emit(b"\\pc")?,
        Charset::Pca => writer.emit(b"\\pca")?,
    }
    writer.emit(b"\\uc1\\deff")?;
    writer.emit_decimal(u32::from(writer.options.default_font))?;
    writer.emit(b" ")?;
    Ok(())
}

fn copy_bytes(output: &mut [u8; 32], bytes: &[u8]) -> usize {
    output[..bytes.len()].copy_from_slice(bytes);
    bytes.len()
}

fn copy_char(output: &mut [u8; 32], character: char) -> usize {
    output[0] = u8::try_from(character as u32).unwrap_or(b'?');
    1
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Document;
    use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits};
    use std::mem::size_of;
    use std::num::{NonZeroU64, NonZeroUsize};

    fn configured_context(
        memory: u64,
        input: u64,
        output: u64,
        objects: u64,
        work: u64,
    ) -> ExecutionContext {
        let budget = Budget::root(
            "rtf-stream-test",
            Limits::new(memory, input, output, objects, 64, work),
        );
        let (_source, token) = CancellationSource::pair();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroU64::new(1024 * 1024).expect("nonzero"),
            1,
        )
        .expect("valid limits");
        ExecutionContext::new(budget, token, execution)
    }

    fn context(output: u64) -> ExecutionContext {
        configured_context(1024 * 1024, output, output, 10_000, 10_000_000)
    }

    fn hierarchical_context(
        parent_limits: Limits,
        child_limits: Limits,
    ) -> (Budget, ExecutionContext) {
        let parent = Budget::root("rtf-stream-parent", parent_limits);
        let child = parent.child("rtf-stream-child", child_limits);
        let (_source, token) = CancellationSource::pair();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroU64::new(1024 * 1024).expect("nonzero"),
            1,
        )
        .expect("valid limits");
        (parent, ExecutionContext::new(child, token, execution))
    }

    fn writer() -> StreamingRtfWriter<Vec<u8>> {
        StreamingRtfWriter::new(
            Vec::new(),
            context(16 * 1024 * 1024),
            StreamingRtfLimits::default(),
        )
        .expect("writer")
    }

    fn finish_one_run<W: Write>(writer: &mut StreamingRtfWriter<W>, text: &str) {
        writer.start_paragraph().expect("paragraph");
        writer.start_run().expect("run");
        writer.write_all(text.as_bytes()).expect("text");
        writer.finish_run().expect("run finish");
        writer.finish_paragraph().expect("paragraph finish");
    }

    #[derive(Debug, Default)]
    struct InstrumentedSink {
        bytes: Vec<u8>,
        write_calls: usize,
        max_request: usize,
    }

    impl Write for InstrumentedSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            self.write_calls = self.write_calls.saturating_add(1);
            self.max_request = self.max_request.max(bytes.len());
            self.bytes.extend_from_slice(bytes);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    fn legacy_write_calls(text: &str) -> usize {
        // Header emission for the default Windows-1252/\deff0 options.
        let mut calls = 9usize;
        // start_paragraph, start_run, finish_run, and finish_paragraph.
        calls = calls.saturating_add(4);
        let mut pending_cr = false;
        for character in text.chars() {
            if pending_cr {
                calls = calls.saturating_add(1);
                if character == '\n' {
                    pending_cr = false;
                    continue;
                }
                pending_cr = false;
            }
            if character == '\r' {
                pending_cr = true;
            } else {
                calls = calls.saturating_add(1);
            }
        }
        if pending_cr {
            calls = calls.saturating_add(1);
        }
        // finish() emits the document's final brace.
        calls.saturating_add(1)
    }

    #[test]
    fn contiguous_ascii_spans_reduce_calls_and_reopen_long_text() {
        let inputs = [
            "plain ASCII span ".repeat(256),
            "plain \\ braces { and } plus café ".repeat(128),
            "café ".repeat(256),
        ];
        for text in inputs {
            let input_bytes = u64::try_from(text.len()).expect("input length");
            let limits = StreamingRtfLimits::new(input_bytes, 1_048_576, 1, 1, 64);
            let context = configured_context(
                1_048_576,
                input_bytes,
                1_048_576,
                16,
                input_bytes.saturating_add(64),
            );
            let mut value = StreamingRtfWriter::new(InstrumentedSink::default(), context, limits)
                .expect("writer");
            finish_one_run(&mut value, &text);
            let sink = value.finish().expect("finish");
            let legacy_calls = legacy_write_calls(&text);
            assert!(
                sink.write_calls.saturating_mul(4) < legacy_calls.saturating_mul(3),
                "batched calls {} were not materially below legacy {}",
                sink.write_calls,
                legacy_calls
            );

            assert!(sink.max_request <= MAX_ASCII_BATCH_BYTES);
            let document = Document::from_bytes(&sink.bytes).expect("reopen");
            assert_eq!(
                document
                    .body()
                    .paragraphs()
                    .next()
                    .expect("paragraph")
                    .to_text(),
                text
            );
        }
    }

    #[test]
    fn batching_preserves_prechange_wire_for_representative_text() {
        let cases = [
            (
                "ASCII",
                br"{\rtf1\ansi\ansicpg1252\uc1\deff0 \pard {\plain ASCII}\par }".as_slice(),
            ),
            (
                r"slashes \ braces { }",
                br"{\rtf1\ansi\ansicpg1252\uc1\deff0 \pard {\plain slashes \\ braces \{ \}}\par }"
                    .as_slice(),
            ),
            (
                "CP1252 é €",
                br"{\rtf1\ansi\ansicpg1252\uc1\deff0 \pard {\plain CP1252 \'e9 \'80}\par }"
                    .as_slice(),
            ),
            (
                "line\r\nnext\rlast",
                br"{\rtf1\ansi\ansicpg1252\uc1\deff0 \pard {\plain line\line next\line last}\par }"
                    .as_slice(),
            ),
        ];
        for (text, expected) in cases {
            let input_bytes = u64::try_from(text.len()).expect("input length");
            let output_bytes = u64::try_from(expected.len()).expect("output length");
            let limits = StreamingRtfLimits::new(input_bytes, output_bytes, 1, 1, 64);
            let context = configured_context(
                1_048_576,
                input_bytes,
                output_bytes,
                16,
                input_bytes.saturating_add(64),
            );
            let mut value = StreamingRtfWriter::new(InstrumentedSink::default(), context, limits)
                .expect("writer");
            finish_one_run(&mut value, text);
            let sink = value.finish().expect("finish");
            assert_eq!(sink.bytes, expected);
        }
    }

    #[test]
    fn hierarchical_work_one_under_falls_back_to_scalar_progress() {
        let span = b"abcdefgh";
        let (parent, context) = hierarchical_context(
            Limits::new(1_048_576, 1024, 1024, 16, 64, 9),
            Limits::new(1_048_576, 1024, 1024, 16, 64, 1024),
        );
        let limits = StreamingRtfLimits::new(span.len() as u64, 1024, 1, 1, 64);
        let mut value = StreamingRtfWriter::new(Vec::new(), context, limits).expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        let prefix = value.output_bytes();
        let error = value.write_utf8(span).expect_err("one-under parent Work");
        assert!(matches!(
            error,
            StreamingRtfError::LimitExceeded {
                resource: "work",
                observed: 10,
                limit: 9,
                written,
            } if written == prefix + 7
        ));
        assert_eq!(value.output_bytes(), prefix + 7);
        assert_eq!(value.input_bytes(), span.len() as u64);
        assert_eq!(parent.used(Resource::Work), 9);
        assert!(value.is_poisoned());
    }

    #[test]
    fn hierarchical_output_one_under_falls_back_to_scalar_progress() {
        let span = b"abcdefgh";
        let prefix = u64::try_from(
            b"{\\rtf1\\ansi\\ansicpg1252\\uc1\\deff0 ".len()
                + b"\\pard ".len()
                + b"{\\plain ".len(),
        )
        .expect("prefix length");
        let output_limit = prefix + span.len() as u64 - 1;
        let (parent, context) = hierarchical_context(
            Limits::new(1_048_576, 1024, output_limit, 16, 64, 1024),
            Limits::new(1_048_576, 1024, 1024, 16, 64, 1024),
        );
        let limits = StreamingRtfLimits::new(span.len() as u64, 1024, 1, 1, 64);
        let mut value = StreamingRtfWriter::new(Vec::new(), context, limits).expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        assert_eq!(value.output_bytes(), prefix);
        let error = value.write_utf8(span).expect_err("one-under parent output");
        assert!(matches!(
            error,
            StreamingRtfError::LimitExceeded {
                resource: "output bytes",
                observed,
                limit,
                written,
            } if observed == output_limit + 1 && limit == output_limit && written == output_limit
        ));
        assert_eq!(value.output_bytes(), output_limit);
        assert_eq!(value.input_bytes(), span.len() as u64);
        assert_eq!(parent.used(Resource::OutputBytes), output_limit);
        assert!(value.is_poisoned());
    }

    #[test]
    fn empty_document_is_deterministic_and_reopens() {
        let first = writer().finish().expect("finish");
        let second = writer().finish().expect("finish");
        assert_eq!(first, second);
        assert_eq!(first, b"{\\rtf1\\ansi\\ansicpg1252\\uc1\\deff0 }");
        let document = Document::from_bytes(&first).expect("parse");
        assert!(document.is_empty());
    }

    #[test]
    fn paragraphs_and_runs_are_explicitly_ordered() {
        let mut value = writer();
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        value.write_all(b"first").expect("text");
        value.finish_run().expect("run");
        value.start_run().expect("second run");
        value.write_all(b" second").expect("text");
        value.finish_run().expect("run");
        value.finish_paragraph().expect("paragraph");
        value.start_paragraph().expect("second paragraph");
        value.finish_paragraph().expect("empty paragraph");
        let bytes = value.finish().expect("finish");
        let document = Document::from_bytes(&bytes).expect("parse");
        let paragraphs: Vec<String> = document
            .body()
            .paragraphs()
            .map(|paragraph| paragraph.to_text())
            .collect();
        assert_eq!(paragraphs, ["first second", ""]);
        assert_eq!(document.paragraph_count(), 2);
    }

    #[test]
    fn escaping_and_code_page_unicode_are_round_trippable() {
        let mut value = writer();
        finish_one_run(
            &mut value,
            "slashes \\ braces { } \0\u{1f}\u{7f} é 你 😀\n\t",
        );
        let bytes = value.finish().expect("finish");
        let text = String::from_utf8(bytes.clone()).expect("wire UTF-8");
        assert!(text.contains(r"\'00"));
        assert!(text.contains(r"\'1f"));
        assert!(text.contains(r"\'7f"));
        assert!(text.contains(r"\'e9"));
        assert!(text.contains(r"\u"));
        assert!(text.contains(r"\line "));
        assert!(text.contains(r"\tab "));
        let document = Document::from_bytes(&bytes).expect("parse");
        assert_eq!(
            document
                .body()
                .paragraphs()
                .next()
                .expect("paragraph")
                .to_text(),
            "slashes \\ braces { } \0\u{1f}\u{7f} é 你 😀\n\t"
        );
    }

    #[test]
    fn crlf_is_one_line_break_even_when_split_across_writes() {
        let mut value = writer();
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        value.write_all(b"before\r").expect("carriage return");
        value.write_all(b"\nafter").expect("line feed");
        value.finish_run().expect("run");
        value.finish_paragraph().expect("paragraph");
        let bytes = value.finish().expect("finish");
        let wire = String::from_utf8(bytes.clone()).expect("wire");
        assert_eq!(wire.matches(r"\line ").count(), 1);
        let document = Document::from_bytes(&bytes).expect("parse");
        assert_eq!(
            document
                .body()
                .paragraphs()
                .next()
                .expect("paragraph")
                .to_text(),
            "before\nafter"
        );
    }

    #[test]
    fn limits_are_exact_and_one_over() {
        let limits = StreamingRtfLimits::new(3, 1024, 1, 1, 64);
        let mut value = StreamingRtfWriter::new(Vec::new(), context(4096), limits).expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        assert_eq!(value.write(b"abc").expect("exact input"), 3);
        assert!(value.write(b"d").is_err());

        let mut value = StreamingRtfWriter::new(Vec::new(), context(4096), limits).expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        assert!(value.write_all(b"a").is_ok());
        assert!(value.start_paragraph().is_err());
        assert!(!value.is_poisoned(), "state errors remain retryable");

        let paragraph_limits = StreamingRtfLimits {
            max_paragraphs: 1,
            ..StreamingRtfLimits::default()
        };
        let mut value =
            StreamingRtfWriter::new(Vec::new(), context(4096), paragraph_limits).expect("writer");
        value.start_paragraph().expect("exact paragraph");
        value.finish_paragraph().expect("paragraph");
        assert!(value.start_paragraph().is_err());

        let run_limits = StreamingRtfLimits {
            max_runs: 1,
            ..StreamingRtfLimits::default()
        };
        let mut value =
            StreamingRtfWriter::new(Vec::new(), context(4096), run_limits).expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("exact run");
        value.finish_run().expect("run");
        assert!(value.start_run().is_err());
    }

    #[test]
    fn fixed_memory_and_budget_dimensions_are_exact_and_one_under() {
        let under = configured_context(FIXED_SCRATCH_BYTES - 1, 4096, 4096, 16, 16);
        let error = StreamingRtfWriter::new(Vec::new(), under, StreamingRtfLimits::default())
            .err()
            .expect("one-under fixed scratch");
        assert!(matches!(
            error,
            StreamingRtfError::LimitExceeded {
                resource: "memory",
                ..
            }
        ));

        let exact = configured_context(FIXED_SCRATCH_BYTES, 4096, 4096, 16, 16);
        StreamingRtfWriter::new(Vec::new(), exact, StreamingRtfLimits::default())
            .expect("exact fixed scratch");

        let baseline = writer();
        let header_bytes = baseline.output_bytes();
        drop(baseline);
        let output_limits = StreamingRtfLimits {
            max_output_bytes: header_bytes,
            ..StreamingRtfLimits::default()
        };
        StreamingRtfWriter::new(
            Vec::new(),
            configured_context(FIXED_SCRATCH_BYTES, 4096, header_bytes, 16, 16),
            output_limits,
        )
        .expect("exact header output")
        .finish()
        .expect_err("closing brace is one over");
        let output_under = StreamingRtfLimits {
            max_output_bytes: header_bytes.saturating_sub(1),
            ..StreamingRtfLimits::default()
        };
        assert!(
            StreamingRtfWriter::new(
                Vec::new(),
                configured_context(FIXED_SCRATCH_BYTES, 4096, header_bytes, 16, 16),
                output_under,
            )
            .is_err()
        );

        let mut value = StreamingRtfWriter::new(
            Vec::new(),
            configured_context(FIXED_SCRATCH_BYTES, 4096, 4096, 0, 16),
            StreamingRtfLimits::default(),
        )
        .expect("writer");
        assert!(value.start_paragraph().is_err());
        assert!(value.is_poisoned());

        let mut value = StreamingRtfWriter::new(
            Vec::new(),
            configured_context(FIXED_SCRATCH_BYTES, 4096, 4096, 16, 0),
            StreamingRtfLimits::default(),
        )
        .expect("writer");
        assert!(value.start_paragraph().is_err());
        assert!(value.is_poisoned());

        let input_limits = StreamingRtfLimits {
            max_input_bytes: 1,
            ..StreamingRtfLimits::default()
        };
        let mut value =
            StreamingRtfWriter::new(Vec::new(), context(4096), input_limits).expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        assert_eq!(value.write(b"x").expect("exact input"), 1);
        assert!(value.write(b"y").is_err());
        assert!(value.is_poisoned());
    }

    #[derive(Debug, Default)]
    struct CountingSink {
        accepted: u64,
        hash: u64,
    }

    impl Write for CountingSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            for byte in bytes {
                self.hash = self
                    .hash
                    .wrapping_mul(1_099_511_628_211)
                    .wrapping_add(u64::from(*byte) + 1);
            }
            self.accepted = self
                .accepted
                .saturating_add(u64::try_from(bytes.len()).unwrap_or(u64::MAX));
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    fn counted_writer() -> StreamingRtfWriter<CountingSink> {
        let limits = StreamingRtfLimits {
            max_input_bytes: 2_000_000,
            ..StreamingRtfLimits::default()
        };
        StreamingRtfWriter::new(CountingSink::default(), context(2_000_000), limits)
            .expect("writer")
    }

    fn write_counted_run(repetitions: usize) -> CountingSink {
        let mut value = counted_writer();
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        for _ in 0..repetitions {
            value.write_all(b"x").expect("text");
        }
        value.finish_run().expect("run");
        value.finish_paragraph().expect("paragraph");
        value.finish().expect("finish")
    }

    #[test]
    fn non_seek_and_large_chunked_input_use_one_bounded_state() {
        let small = write_counted_run(1);
        let large = write_counted_run(1_000_000);
        assert!(large.accepted > small.accepted);
        assert_ne!(large.hash, small.hash);

        let mut small_state = counted_writer();
        let mut large_state = counted_writer();
        small_state.start_paragraph().expect("small paragraph");
        small_state.start_run().expect("small run");
        large_state.start_paragraph().expect("large paragraph");
        large_state.start_run().expect("large run");
        small_state.write_all(b"x").expect("small text");
        for _ in 0..1_000_000 {
            large_state.write_all(b"x").expect("large text");
        }
        assert_eq!(size_of_val(&small_state), size_of_val(&large_state));
        assert_eq!(
            size_of_val(&small_state),
            size_of::<StreamingRtfWriter<CountingSink>>()
        );
        assert_eq!(
            small_state.pending_utf8.len(),
            large_state.pending_utf8.len()
        );
        assert_eq!(small_state.pending_len, large_state.pending_len);
        assert_eq!(small_state.pending_cr, large_state.pending_cr);
        assert_eq!(small_state.input_bytes(), 1);
        assert_eq!(large_state.input_bytes(), 1_000_000);
        assert_eq!(size_of::<[u8; 4]>(), 4);
        assert_eq!(size_of::<[u8; 32]>(), 32);
    }

    #[test]
    fn counting_sink_is_forward_only() {
        struct NonSeek(CountingSink);
        impl Write for NonSeek {
            fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
                self.0.write(bytes)
            }
            fn flush(&mut self) -> io::Result<()> {
                self.0.flush()
            }
        }

        let limits = StreamingRtfLimits {
            max_input_bytes: 1_000_000,
            ..StreamingRtfLimits::default()
        };
        let mut value =
            StreamingRtfWriter::new(NonSeek(CountingSink::default()), context(2_000_000), limits)
                .expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        for _ in 0..100_000 {
            value.write_all(b"x").expect("chunk");
        }
        value.finish_run().expect("run");
        value.finish_paragraph().expect("paragraph");
        let sink = value.finish().expect("finish");
        assert!(sink.0.accepted > 100_000);
    }

    #[test]
    fn split_utf8_and_sink_progress_are_handled() {
        let mut value = writer();
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        assert_eq!(value.write(&[0xe2]).expect("prefix"), 1);
        assert_eq!(value.write(&[0x82, 0xac]).expect("suffix"), 2);
        value.finish_run().expect("run");
        value.finish_paragraph().expect("paragraph");
        let bytes = value.finish().expect("finish");
        let document = Document::from_bytes(&bytes).expect("parse");
        assert_eq!(
            document
                .body()
                .paragraphs()
                .next()
                .expect("paragraph")
                .to_text(),
            "€"
        );

        #[derive(Debug)]
        struct Partial {
            bytes: Vec<u8>,
            remaining: usize,
        }
        impl Write for Partial {
            fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
                if self.remaining == 0 {
                    return Err(io::Error::new(io::ErrorKind::BrokenPipe, "closed"));
                }
                let count = bytes.len().min(self.remaining).min(2);
                self.bytes.extend_from_slice(&bytes[..count]);
                self.remaining -= count;
                Ok(count)
            }
            fn flush(&mut self) -> io::Result<()> {
                Ok(())
            }
        }
        let mut failing = StreamingRtfWriter::new(
            Partial {
                bytes: Vec::new(),
                remaining: 60,
            },
            context(4096),
            StreamingRtfLimits::default(),
        )
        .expect("writer");
        failing.start_paragraph().expect("paragraph");
        failing.start_run().expect("run");
        let error = failing
            .write(b"this eventually fails")
            .expect_err("sink failure");
        assert_eq!(error.kind(), io::ErrorKind::BrokenPipe);
        assert!(failing.output_bytes() > 0);
    }

    #[test]
    fn malformed_utf8_preserves_ascii_prefix_suffix_and_split_progress() {
        let mut value = writer();
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        let output_prefix = value.output_bytes();
        let ascii_prefix = b"ascii-prefix";
        let error = value
            .write_utf8(b"ascii-prefix\xffascii-suffix")
            .expect_err("malformed suffix");
        assert!(matches!(
            error,
            StreamingRtfError::InvalidInput { written, .. } if written
                == output_prefix + u64::try_from(ascii_prefix.len()).expect("prefix length")
        ));
        assert_eq!(
            value.input_bytes(),
            u64::try_from(ascii_prefix.len()).expect("prefix length")
        );
        assert_eq!(
            value.output_bytes(),
            output_prefix + u64::try_from(ascii_prefix.len()).expect("prefix length")
        );
        assert!(value.is_poisoned());

        let mut value = writer();
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        let mut first = b"ascii-prefix".to_vec();
        first.push(0xe2);
        assert_eq!(
            value.write_utf8(&first).expect("split UTF-8 prefix"),
            first.len()
        );
        assert_eq!(
            value
                .write_utf8(b"\x82\xacascii-suffix")
                .expect("split UTF-8 suffix"),
            2 + b"ascii-suffix".len()
        );
        value.finish_run().expect("run");
        value.finish_paragraph().expect("paragraph");
        let bytes = value.finish().expect("finish");
        let document = Document::from_bytes(&bytes).expect("parse");
        assert_eq!(
            document
                .body()
                .paragraphs()
                .next()
                .expect("paragraph")
                .to_text(),
            "ascii-prefix€ascii-suffix"
        );

        let mut value = writer();
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        let mut split_prefix = b"ascii-prefix".to_vec();
        split_prefix.push(0xe2);
        assert_eq!(
            value
                .write_utf8(&split_prefix)
                .expect("malformed split prefix"),
            split_prefix.len()
        );
        let error = value
            .write_utf8(b"\x28suffix")
            .expect_err("malformed split continuation");
        let accepted = split_prefix.len().saturating_add(2);
        assert!(matches!(error, StreamingRtfError::InvalidInput { .. }));
        assert_eq!(
            value.input_bytes(),
            u64::try_from(accepted).expect("accepted malformed prefix")
        );
        assert!(value.is_poisoned());
    }

    #[test]
    fn incomplete_utf8_cannot_close_a_run() {
        let mut value = writer();
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        assert_eq!(value.write(&[0xe2]).expect("prefix"), 1);
        let error = value.finish_run().expect_err("incomplete run");
        assert!(matches!(error, StreamingRtfError::InvalidInput { .. }));
        assert!(value.is_poisoned());
    }

    #[derive(Debug)]
    struct FaultSink {
        accepted: u64,
        fail_after: Option<u64>,
        max_chunk: usize,
        interrupt_once: bool,
        zero_after: bool,
        flush_error: bool,
    }

    impl Write for FaultSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if self.interrupt_once {
                self.interrupt_once = false;
                return Err(io::Error::new(io::ErrorKind::Interrupted, "retry"));
            }
            if self.fail_after.is_some_and(|limit| self.accepted >= limit) {
                if self.zero_after {
                    return Ok(0);
                }
                return Err(io::Error::new(io::ErrorKind::BrokenPipe, "closed"));
            }
            let available = self
                .fail_after
                .map_or(u64::MAX, |limit| limit.saturating_sub(self.accepted));
            let count = bytes
                .len()
                .min(self.max_chunk.max(1))
                .min(usize::try_from(available).unwrap_or(usize::MAX));
            if count == 0 {
                return Ok(0);
            }
            self.accepted = self
                .accepted
                .saturating_add(u64::try_from(count).unwrap_or(u64::MAX));
            Ok(count)
        }

        fn flush(&mut self) -> io::Result<()> {
            if self.flush_error {
                Err(io::Error::new(io::ErrorKind::BrokenPipe, "flush failed"))
            } else {
                Ok(())
            }
        }
    }

    #[derive(Debug)]
    struct CancellingBatchSink {
        source: CancellationSource,
        armed: bool,
        accepted: u64,
    }

    impl Write for CancellingBatchSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if self.armed {
                self.armed = false;
                self.source.cancel();
                let count = bytes.len().min(2);
                self.accepted = self
                    .accepted
                    .saturating_add(u64::try_from(count).unwrap_or(u64::MAX));
                return Ok(count);
            }
            if self.source.is_cancelled() {
                return Err(io::Error::new(io::ErrorKind::Interrupted, "cancelled"));
            }
            self.accepted = self
                .accepted
                .saturating_add(u64::try_from(bytes.len()).unwrap_or(u64::MAX));
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[derive(Debug)]
    struct FullWriteCancellingSink {
        source: CancellationSource,
        armed: bool,
        accepted: u64,
    }

    impl Write for FullWriteCancellingSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if self.armed {
                self.armed = false;
                self.source.cancel();
                self.accepted = self
                    .accepted
                    .saturating_add(u64::try_from(bytes.len()).unwrap_or(u64::MAX));
                return Ok(bytes.len());
            }
            if self.source.is_cancelled() {
                return Err(io::Error::new(io::ErrorKind::Interrupted, "cancelled"));
            }
            self.accepted = self
                .accepted
                .saturating_add(u64::try_from(bytes.len()).unwrap_or(u64::MAX));
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[derive(Debug, Default)]
    struct ExcessProgressBatchSink {
        armed: bool,
    }

    impl Write for ExcessProgressBatchSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if self.armed {
                return Ok(bytes.len().saturating_add(1));
            }
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    fn default_body_prefix_bytes() -> u64 {
        u64::try_from(
            b"{\\rtf1\\ansi\\ansicpg1252\\uc1\\deff0 ".len()
                + b"\\pard ".len()
                + b"{\\plain ".len(),
        )
        .expect("prefix length")
    }

    fn batched_limits() -> StreamingRtfLimits {
        StreamingRtfLimits::new(8, 1024, 1, 1, 64)
    }

    fn batched_context() -> ExecutionContext {
        configured_context(1_048_576, 8, 1024, 16, 1024)
    }

    #[test]
    fn batched_ascii_partial_sink_commits_exact_prefix_progress() {
        let prefix = default_body_prefix_bytes();
        let sink = FaultSink {
            accepted: 0,
            fail_after: Some(prefix + 3),
            max_chunk: 2,
            interrupt_once: false,
            zero_after: false,
            flush_error: false,
        };
        let mut value =
            StreamingRtfWriter::new(sink, batched_context(), batched_limits()).expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        let error = value
            .write_utf8(b"abcdefgh")
            .expect_err("partial batched sink");
        assert!(matches!(
            error,
            StreamingRtfError::Io {
                kind: io::ErrorKind::BrokenPipe,
                written,
                ..
            } if written == prefix + 3
        ));
        assert_eq!(value.output_bytes(), prefix + 3);
        assert_eq!(value.input_bytes(), 4);
        assert_eq!(
            value.context.budget().used(Resource::OutputBytes),
            prefix + 3
        );
        assert_eq!(value.context.budget().used(Resource::Work), 2 + 4);
        assert!(value.is_poisoned());
    }

    #[test]
    fn batched_ascii_write_zero_commits_only_attempted_prefix() {
        let prefix = default_body_prefix_bytes();
        let sink = FaultSink {
            accepted: 0,
            fail_after: Some(prefix),
            max_chunk: usize::MAX,
            interrupt_once: false,
            zero_after: true,
            flush_error: false,
        };
        let mut value =
            StreamingRtfWriter::new(sink, batched_context(), batched_limits()).expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        let error = value
            .write_utf8(b"abcdefgh")
            .expect_err("zero batched sink");
        assert!(matches!(
            error,
            StreamingRtfError::Io {
                kind: io::ErrorKind::WriteZero,
                written,
                ..
            } if written == prefix
        ));
        assert_eq!(value.output_bytes(), prefix);
        assert_eq!(value.input_bytes(), 1);
        assert_eq!(value.context.budget().used(Resource::OutputBytes), prefix);
        assert_eq!(value.context.budget().used(Resource::Work), 2 + 1);
        assert!(value.is_poisoned());
    }

    #[test]
    fn batched_ascii_cancellation_between_short_writes_commits_exact_prefix() {
        let budget = Budget::root(
            "rtf-stream-batched-cancel-test",
            Limits::new(1_048_576, 1024, 1024, 16, 64, 1024),
        );
        let (source, token) = CancellationSource::pair();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroU64::new(1024).expect("nonzero"),
            1,
        )
        .expect("valid limits");
        let context = ExecutionContext::new(budget, token, execution);
        let sink = CancellingBatchSink {
            source,
            armed: false,
            accepted: 0,
        };
        let mut value = StreamingRtfWriter::new(sink, context, batched_limits()).expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        value.writer.armed = true;
        let prefix = value.output_bytes();
        let error = value
            .write_utf8(b"abcdefgh")
            .expect_err("cancellation between short writes");
        assert!(matches!(
            error,
            StreamingRtfError::Cancelled { written } if written == prefix + 2
        ));
        assert_eq!(value.output_bytes(), prefix + 2);
        assert_eq!(value.input_bytes(), 3);
        assert_eq!(
            value.context.budget().used(Resource::OutputBytes),
            prefix + 2
        );
        assert_eq!(value.context.budget().used(Resource::Work), 2 + 3);
        assert!(value.is_poisoned());
    }

    #[test]
    fn batched_ascii_full_write_cancellation_commits_one_bounded_chunk() {
        let budget = Budget::root(
            "rtf-stream-batched-full-cancel-test",
            Limits::new(1_048_576, 1024, 1024, 16, 64, 1024),
        );
        let (source, token) = CancellationSource::pair();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroU64::new(1024).expect("nonzero"),
            1,
        )
        .expect("valid limits");
        let context = ExecutionContext::new(budget, token, execution);
        let sink = FullWriteCancellingSink {
            source,
            armed: false,
            accepted: 0,
        };
        let limits = StreamingRtfLimits::new(64, 1024, 1, 1, 64);
        let mut value = StreamingRtfWriter::new(sink, context, limits).expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        value.writer.armed = true;
        let prefix = value.output_bytes();
        let error = value
            .write_utf8(&[b'a'; 64])
            .expect_err("full-write cancellation");
        let chunk = u64::try_from(MAX_ASCII_BATCH_BYTES).expect("batch length");
        assert!(matches!(
            error,
            StreamingRtfError::Cancelled { written } if written == prefix + chunk
        ));
        assert_eq!(value.output_bytes(), prefix + chunk);
        assert_eq!(value.input_bytes(), chunk);
        assert_eq!(
            value.context.budget().used(Resource::OutputBytes),
            prefix + chunk
        );
        assert_eq!(value.context.budget().used(Resource::Work), 2 + chunk);
        assert_eq!(value.writer.accepted, prefix + chunk);
        assert!(value.is_poisoned());
    }

    #[test]
    fn batched_ascii_excess_progress_poison_reports_attempted_prefix() {
        let mut value = StreamingRtfWriter::new(
            ExcessProgressBatchSink::default(),
            batched_context(),
            batched_limits(),
        )
        .expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        value.writer.armed = true;
        let prefix = value.output_bytes();
        let error = value
            .write_utf8(b"abcdefgh")
            .expect_err("excess batch progress");
        assert!(matches!(
            error,
            StreamingRtfError::InvalidInput { written, .. } if written == prefix
        ));
        assert_eq!(value.output_bytes(), prefix);
        assert_eq!(value.input_bytes(), 1);
        assert_eq!(value.context.budget().used(Resource::OutputBytes), prefix);
        assert_eq!(value.context.budget().used(Resource::Work), 2 + 1);
        assert!(value.is_poisoned());
    }

    #[test]
    fn short_writes_and_interrupted_progress_finish_exactly() {
        let sink = FaultSink {
            accepted: 0,
            fail_after: None,
            max_chunk: 1,
            interrupt_once: true,
            zero_after: false,
            flush_error: false,
        };
        let mut value = StreamingRtfWriter::new(sink, context(4096), StreamingRtfLimits::default())
            .expect("writer");
        finish_one_run(&mut value, "short");
        let sink = value.finish().expect("finish");
        assert!(sink.accepted > 0);
    }

    #[test]
    fn write_zero_poison_reports_exact_progress() {
        let sink = FaultSink {
            accepted: 0,
            fail_after: Some(100),
            max_chunk: usize::MAX,
            interrupt_once: false,
            zero_after: true,
            flush_error: false,
        };
        let mut value = StreamingRtfWriter::new(sink, context(4096), StreamingRtfLimits::default())
            .expect("writer");
        value.start_paragraph().expect("paragraph");
        value.start_run().expect("run");
        let error = value.write(&[b'x'; 128]).expect_err("write zero");
        assert_eq!(error.kind(), io::ErrorKind::WriteZero);
        assert_eq!(value.output_bytes(), 100);
        assert!(value.is_poisoned());
    }

    #[test]
    fn flush_failure_poison_and_final_brace_failure_preserve_progress() {
        let mut flush_writer = StreamingRtfWriter::new(
            FaultSink {
                accepted: 0,
                fail_after: None,
                max_chunk: usize::MAX,
                interrupt_once: false,
                zero_after: false,
                flush_error: true,
            },
            context(4096),
            StreamingRtfLimits::default(),
        )
        .expect("writer");
        finish_one_run(&mut flush_writer, "flush");
        let flush_error = flush_writer.flush().expect_err("flush failure");
        assert_eq!(flush_error.kind(), io::ErrorKind::BrokenPipe);
        assert!(flush_writer.is_poisoned());

        let mut prefix = writer();
        finish_one_run(&mut prefix, "final");
        let prefix_bytes = prefix.output_bytes();
        let mut final_writer = StreamingRtfWriter::new(
            FaultSink {
                accepted: 0,
                fail_after: Some(prefix_bytes),
                max_chunk: usize::MAX,
                interrupt_once: false,
                zero_after: true,
                flush_error: false,
            },
            context(4096),
            StreamingRtfLimits::default(),
        )
        .expect("writer");
        finish_one_run(&mut final_writer, "final");
        let error = final_writer.finish().expect_err("final brace failure");
        assert!(matches!(
            error,
            StreamingRtfError::Io {
                kind: io::ErrorKind::WriteZero,
                written,
                ..
            } if written == prefix_bytes
        ));
    }

    #[test]
    fn cancellation_is_observed_before_next_output() {
        let budget = Budget::root(
            "rtf-stream-cancel-test",
            Limits::new(1024, 1024, 1024, 16, 64, 1024),
        );
        let (source, token) = CancellationSource::pair();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroU64::new(1024).expect("nonzero"),
            1,
        )
        .expect("valid limits");
        let context = ExecutionContext::new(budget, token, execution);
        let mut value = StreamingRtfWriter::new(Vec::new(), context, StreamingRtfLimits::default())
            .expect("writer");
        source.cancel();
        let error = value.start_paragraph().expect_err("cancelled");
        assert!(matches!(error, StreamingRtfError::Cancelled { .. }));

        let budget = Budget::root(
            "rtf-stream-flush-cancel-test",
            Limits::new(1024, 1024, 1024, 16, 64, 1024),
        );
        let (source, token) = CancellationSource::pair();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroU64::new(1024).expect("nonzero"),
            1,
        )
        .expect("valid limits");
        let context = ExecutionContext::new(budget, token, execution);
        let mut flush_value =
            StreamingRtfWriter::new(Vec::new(), context, StreamingRtfLimits::default())
                .expect("writer");
        source.cancel();
        let flush_error = flush_value.flush().expect_err("cancelled flush");
        assert_eq!(flush_error.kind(), io::ErrorKind::Interrupted);
        assert!(flush_value.is_poisoned());
    }

    #[test]
    fn cancellation_is_checked_between_short_sink_retries() {
        use std::sync::{
            Arc,
            atomic::{AtomicBool, AtomicUsize, Ordering},
        };

        #[derive(Debug)]
        struct CancellingSink {
            source: CancellationSource,
            armed: Arc<AtomicBool>,
            calls: Arc<AtomicUsize>,
            interrupted_after_cancel: u8,
        }

        impl Write for CancellingSink {
            fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
                self.calls.fetch_add(1, Ordering::SeqCst);
                if self.armed.swap(false, Ordering::SeqCst) {
                    self.source.cancel();
                    return Ok(bytes.len().min(1));
                }
                if !self.source.is_cancelled() {
                    return Ok(bytes.len());
                }
                if self.source.is_cancelled() && self.interrupted_after_cancel < 4 {
                    self.interrupted_after_cancel = self.interrupted_after_cancel.saturating_add(1);
                    return Err(io::Error::new(io::ErrorKind::Interrupted, "retry"));
                }
                Err(io::Error::new(
                    io::ErrorKind::BrokenPipe,
                    "cancelled sink retried",
                ))
            }

            fn flush(&mut self) -> io::Result<()> {
                Ok(())
            }
        }

        let budget = Budget::root(
            "rtf-stream-cancel-retry-test",
            Limits::new(1024, 1024, 1024, 16, 64, 1024),
        );
        let (source, token) = CancellationSource::pair();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroUsize::new(1).expect("nonzero"),
            NonZeroU64::new(1024).expect("nonzero"),
            1,
        )
        .expect("valid limits");
        let context = ExecutionContext::new(budget, token, execution);
        let armed = Arc::new(AtomicBool::new(false));
        let calls = Arc::new(AtomicUsize::new(0));
        let mut value = StreamingRtfWriter::new(
            CancellingSink {
                source: source.clone(),
                armed: Arc::clone(&armed),
                calls: Arc::clone(&calls),
                interrupted_after_cancel: 0,
            },
            context,
            StreamingRtfLimits::default(),
        )
        .expect("writer");
        let header_bytes = value.output_bytes();
        let calls_before_body = calls.load(Ordering::SeqCst);
        armed.store(true, Ordering::SeqCst);

        let error = value
            .start_paragraph()
            .expect_err("cancellation after a short sink write");
        assert!(matches!(
            error,
            StreamingRtfError::Cancelled { written } if written == header_bytes + 1
        ));
        assert_eq!(value.output_bytes(), header_bytes + 1);
        assert_eq!(calls.load(Ordering::SeqCst), calls_before_body + 1);
        assert!(source.is_cancelled());
        assert!(value.is_poisoned());
    }
}
