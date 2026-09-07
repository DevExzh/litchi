//! Bounded, replayable generated-member publication.
//!
//! A replayable member is deliberately separate from [`RegeneratedEntry`].
//! The callback is invoked once to measure its logical output and once to
//! emit it after the complete preservation layout has been preflighted.  The
//! callback receives only a writer: format owners own the fresh, verified
//! source reader used by each invocation.
//!
//! Deflate input is normalized through a fixed 64 KiB window, with a fixed
//! 64 KiB compressor output buffer. Together with preservation's existing
//! 64 KiB source-copy buffer, replay uses at most 192 KiB of fixed scratch,
//! plus the backend compressor state and small digest/framing state. No
//! scratch allocation scales with the configured decoded, compressed, or
//! archive ceilings.

use super::{
    PreparedCentral, PreparedEntry, PreparedLocal, PreservationEntryId, PreservationIndex,
    write_prepared_central, write_prepared_local, write_prepared_tail,
};
use crate::accounting::{AccountingWriteKind, ZipOperationAccounting};
use crate::crc::crc32_chunk;
use crate::path::ZipFilePath;
use crate::writer::prepare_sized_member;
use crate::{CompressionMethod, Error, ErrorKind, ReaderAt};
use flate2::{Compress, Compression, FlushCompress, Status};
use sha2_legacy::{Digest, Sha256};
use std::fmt;
use std::io::{self, Write};

const DEFAULT_MAX_DECODED_BYTES: u64 = 256 * 1024 * 1024;
const DEFAULT_MAX_COMPRESSED_BYTES: u64 = 256 * 1024 * 1024;
const DEFAULT_MAX_ARCHIVE_BYTES: u64 = 2 * 1024 * 1024 * 1024;
const REPLAY_INPUT_WINDOW_SIZE: usize = 64 * 1024;
const REPLAY_COMPRESS_BUFFER_SIZE: usize = 64 * 1024;

/// Finite ceilings for one replayable member publication.
///
/// The decoded and compressed limits bound the generated member. The archive
/// limit includes every preserved local span, generated framing, central
/// directory record, tail, and archive comment. The constructor rejects zero
/// and `u64::MAX`; callers that need a larger operation must choose another
/// explicit finite value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReplayLimits {
    pub max_decoded_bytes: u64,
    pub max_compressed_bytes: u64,
    pub max_archive_bytes: u64,
}

impl ReplayLimits {
    /// The default finite replay envelope.
    pub const DEFAULT: Self = Self {
        max_decoded_bytes: DEFAULT_MAX_DECODED_BYTES,
        max_compressed_bytes: DEFAULT_MAX_COMPRESSED_BYTES,
        max_archive_bytes: DEFAULT_MAX_ARCHIVE_BYTES,
    };

    /// Construct finite replay limits.
    pub fn new(
        max_decoded_bytes: u64,
        max_compressed_bytes: u64,
        max_archive_bytes: u64,
    ) -> Result<Self, Error> {
        let limits = Self {
            max_decoded_bytes,
            max_compressed_bytes,
            max_archive_bytes,
        };
        limits.validate()?;
        Ok(limits)
    }

    fn validate(self) -> Result<(), Error> {
        if self.max_decoded_bytes == 0
            || self.max_compressed_bytes == 0
            || self.max_archive_bytes == 0
            || self.max_decoded_bytes == u64::MAX
            || self.max_compressed_bytes == u64::MAX
            || self.max_archive_bytes == u64::MAX
        {
            return Err(ErrorKind::InvalidInput {
                msg: "replay limits must be finite and nonzero".to_string(),
            }
            .into());
        }
        Ok(())
    }
}

impl Default for ReplayLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Pass in which a replay operation failed or detected a mismatch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReplayPass {
    /// Callback output was measured without touching the publication sink.
    Measure,
    /// The output layout was validated before the callback was emitted.
    Preflight,
    /// Callback output or the remaining archive was emitted to the sink.
    Emit,
}

/// Bounded resource used by a replay limit failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReplayResource {
    DecodedBytes,
    CompressedBytes,
    ArchiveBytes,
}

impl fmt::Display for ReplayResource {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::DecodedBytes => "decoded replay bytes",
            Self::CompressedBytes => "compressed replay bytes",
            Self::ArchiveBytes => "replay archive bytes",
        })
    }
}

/// Truthful progress accepted by the forward-only publication sink.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReplayProgress {
    Untouched,
    Prefix { accepted: u64 },
    CompleteUnflushed { bytes: u64 },
    Complete { bytes: u64 },
    Indeterminate { accepted_before: u64 },
}

impl ReplayProgress {
    /// Bytes definitely accepted by the sink before the failure.
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

/// Result of one successful replay pass.
///
/// Replay identity compares decoded size, decoded CRC-32, decoded SHA-256,
/// compressed size, and compressed SHA-256. The digest check makes accidental
/// drift highly unlikely; as with every cryptographic digest, it is not an
/// unqualified mathematical proof against a deliberate SHA-256 collision.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReplayMeasurement {
    decoded_size: u64,
    compressed_size: u64,
    decoded_crc32: u32,
    decoded_sha256: [u8; 32],
    compressed_sha256: [u8; 32],
}

impl ReplayMeasurement {
    #[must_use]
    pub const fn decoded_size(self) -> u64 {
        self.decoded_size
    }

    #[must_use]
    pub const fn compressed_size(self) -> u64 {
        self.compressed_size
    }

    #[must_use]
    pub const fn decoded_crc32(self) -> u32 {
        self.decoded_crc32
    }

    #[must_use]
    pub const fn decoded_sha256(self) -> [u8; 32] {
        self.decoded_sha256
    }

    #[must_use]
    pub const fn compressed_sha256(self) -> [u8; 32] {
        self.compressed_sha256
    }
}

/// A typed failure from replay measurement or publication.
#[derive(Debug)]
#[non_exhaustive]
pub enum ReplayPublicationError<E> {
    Archive {
        pass: ReplayPass,
        source: Error,
        callback_error: Option<E>,
        progress: ReplayProgress,
    },
    Callback {
        pass: ReplayPass,
        source: E,
        progress: ReplayProgress,
    },
    Limit {
        pass: ReplayPass,
        resource: ReplayResource,
        actual: u64,
        maximum: u64,
        callback_error: Option<E>,
        progress: ReplayProgress,
    },
    Sink {
        pass: ReplayPass,
        source: io::Error,
        callback_error: Option<E>,
        progress: ReplayProgress,
    },
    NonDeterministic {
        expected: Box<ReplayMeasurement>,
        actual: Box<ReplayMeasurement>,
        progress: ReplayProgress,
    },
}

impl<E> ReplayPublicationError<E> {
    #[must_use]
    pub const fn progress(&self) -> ReplayProgress {
        match self {
            Self::Archive { progress, .. }
            | Self::Callback { progress, .. }
            | Self::Limit { progress, .. }
            | Self::Sink { progress, .. }
            | Self::NonDeterministic { progress, .. } => *progress,
        }
    }

    #[must_use]
    pub const fn pass(&self) -> ReplayPass {
        match self {
            Self::Archive { pass, .. }
            | Self::Callback { pass, .. }
            | Self::Limit { pass, .. }
            | Self::Sink { pass, .. } => *pass,
            Self::NonDeterministic { .. } => ReplayPass::Emit,
        }
    }

    #[must_use]
    pub const fn callback(&self) -> Option<&E> {
        match self {
            Self::Archive { callback_error, .. }
            | Self::Limit { callback_error, .. }
            | Self::Sink { callback_error, .. } => callback_error.as_ref(),
            Self::Callback { source, .. } => Some(source),
            Self::NonDeterministic { .. } => None,
        }
    }
}

impl<E: fmt::Display> fmt::Display for ReplayPublicationError<E> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Archive {
                pass,
                source,
                callback_error,
                ..
            } => match callback_error {
                Some(callback) => write!(
                    formatter,
                    "replay archive failure during {pass:?}: {source} (callback also failed: {callback})"
                ),
                None => write!(
                    formatter,
                    "replay archive failure during {pass:?}: {source}"
                ),
            },
            Self::Callback { pass, source, .. } => {
                write!(
                    formatter,
                    "replay callback failed during {pass:?}: {source}"
                )
            },
            Self::Limit {
                pass,
                resource,
                actual,
                maximum,
                callback_error,
                ..
            } => match callback_error {
                Some(callback) => write!(
                    formatter,
                    "replay {resource} limit failed during {pass:?}: {actual} > {maximum} (callback also failed: {callback})"
                ),
                None => write!(
                    formatter,
                    "replay {resource} limit failed during {pass:?}: {actual} > {maximum}"
                ),
            },
            Self::Sink {
                pass,
                source,
                callback_error,
                ..
            } => match callback_error {
                Some(callback) => write!(
                    formatter,
                    "replay sink failure during {pass:?}: {source} (callback also failed: {callback})"
                ),
                None => write!(formatter, "replay sink failure during {pass:?}: {source}"),
            },
            Self::NonDeterministic { .. } => {
                formatter.write_str("replay callback produced different output on emission")
            },
        }
    }
}

impl<E: fmt::Debug + fmt::Display + 'static> std::error::Error for ReplayPublicationError<E> {}

/// Run a bounded callback twice while replacing one existing source member.
impl<'source, R> PreservationIndex<'source, R>
where
    R: ReaderAt,
{
    /// Replace exactly one existing member while copying every other member's
    /// validated local span and central record. The callback is invoked once
    /// for measurement and once for emission; it receives logical member
    /// bytes, not ZIP-compressed bytes. The callback is never retried; a panic
    /// propagates to the caller and does not trigger a third invocation.
    pub fn write_replacing_with_replay<W, F, E>(
        &self,
        target: PreservationEntryId,
        compression: CompressionMethod,
        limits: ReplayLimits,
        sink: W,
        callback: F,
    ) -> Result<W, ReplayPublicationError<E>>
    where
        W: Write,
        F: FnMut(&mut dyn Write) -> Result<(), E>,
    {
        let mut accounting = ZipOperationAccounting::default();
        self.write_replacing_with_replay_with_accounting(
            target,
            compression,
            limits,
            sink,
            &mut accounting,
            callback,
        )
    }

    /// Replay one replacement while recording accepted generated payload
    /// bytes in the supplied operation-local accounting value.
    pub fn write_replacing_with_replay_with_accounting<W, F, E>(
        &self,
        target: PreservationEntryId,
        compression: CompressionMethod,
        limits: ReplayLimits,
        sink: W,
        accounting: &mut ZipOperationAccounting,
        mut callback: F,
    ) -> Result<W, ReplayPublicationError<E>>
    where
        W: Write,
        F: FnMut(&mut dyn Write) -> Result<(), E>,
    {
        if let Err(source) = limits.validate() {
            return Err(ReplayPublicationError::Archive {
                pass: ReplayPass::Preflight,
                source,
                callback_error: None,
                progress: ReplayProgress::Untouched,
            });
        }
        if !matches!(
            compression,
            CompressionMethod::Store | CompressionMethod::Deflate
        ) {
            return Err(ReplayPublicationError::Archive {
                pass: ReplayPass::Preflight,
                source: ErrorKind::UnsupportedCompressionMethod(compression.as_id().as_u16())
                    .into(),
                callback_error: None,
                progress: ReplayProgress::Untouched,
            });
        }

        let target_index = match usize::try_from(target.0) {
            Ok(index) => index,
            Err(_) => {
                return Err(ReplayPublicationError::Archive {
                    pass: ReplayPass::Preflight,
                    source: unsupported("invalid replay target ID"),
                    callback_error: None,
                    progress: ReplayProgress::Untouched,
                });
            },
        };
        let source_entry = match self.entries.get(target_index) {
            Some(entry) if entry.id == target => entry,
            _ => {
                return Err(ReplayPublicationError::Archive {
                    pass: ReplayPass::Preflight,
                    source: unsupported("invalid replay target ID"),
                    callback_error: None,
                    progress: ReplayProgress::Untouched,
                });
            },
        };
        let name = match replay_target_name(source_entry) {
            Ok(name) => name,
            Err(source) => {
                return Err(ReplayPublicationError::Archive {
                    pass: ReplayPass::Preflight,
                    source,
                    callback_error: None,
                    progress: ReplayProgress::Untouched,
                });
            },
        };

        let measurement = measure_callback(&mut callback, compression, limits)?;
        let replay = match prepare_replay_target(name, compression, measurement) {
            Ok(replay) => replay,
            Err(source) => {
                return Err(ReplayPublicationError::Archive {
                    pass: ReplayPass::Preflight,
                    source,
                    callback_error: None,
                    progress: ReplayProgress::Untouched,
                });
            },
        };

        let plan = super::PreservationPlan::copy_all(self);
        let mut prepared = match self.prepare(&plan) {
            Ok(prepared) => prepared,
            Err(source) => {
                return Err(ReplayPublicationError::Archive {
                    pass: ReplayPass::Preflight,
                    source,
                    callback_error: None,
                    progress: ReplayProgress::Untouched,
                });
            },
        };
        prepared[target_index] = replay;
        let layout = match self.validate_output_layout(&mut prepared) {
            Ok(layout) => layout,
            Err(source) => {
                return Err(ReplayPublicationError::Archive {
                    pass: ReplayPass::Preflight,
                    source,
                    callback_error: None,
                    progress: ReplayProgress::Untouched,
                });
            },
        };
        if layout.output_size > limits.max_archive_bytes {
            return Err(ReplayPublicationError::Limit {
                pass: ReplayPass::Preflight,
                resource: ReplayResource::ArchiveBytes,
                actual: layout.output_size,
                maximum: limits.max_archive_bytes,
                callback_error: None,
                progress: ReplayProgress::Untouched,
            });
        }

        let mut sink = ReplaySink::new(sink);
        sink.set_expected(layout.output_size);
        let mut copy_buffer = [0u8; super::COPY_CHUNK_SIZE];
        for &index in &self.local_order {
            if index == target_index {
                let framing = match &prepared[index].local {
                    PreparedLocal::Replay { framing, .. } => framing,
                    _ => {
                        return Err(ReplayPublicationError::Archive {
                            pass: ReplayPass::Preflight,
                            source: unsupported("replay target framing was not prepared"),
                            callback_error: None,
                            progress: ReplayProgress::Untouched,
                        });
                    },
                };
                sink.write_all(framing)
                    .map_err(|source| sink_error(ReplayPass::Emit, source, None, &mut sink))?;
                let actual =
                    emit_callback(&mut callback, compression, limits, &mut sink, accounting)?;
                if actual != measurement {
                    return Err(ReplayPublicationError::NonDeterministic {
                        expected: Box::new(measurement),
                        actual: Box::new(actual),
                        progress: sink.progress(),
                    });
                }
                continue;
            }
            if prepared[index].omitted {
                continue;
            }
            if let Err(source) = write_prepared_local(
                &prepared[index].local,
                prepared[index].generated_payload.as_ref(),
                self.source,
                &mut sink,
                &mut copy_buffer,
                accounting,
            ) {
                return Err(map_emit_error(source, &mut sink));
            }
        }

        for (entry, patch) in prepared
            .iter()
            .filter(|entry| !entry.omitted)
            .zip(layout.central_patches.iter())
        {
            let central = entry.central.bytes(&self.entries);
            let is_unchanged = matches!(&entry.central, PreparedCentral::Copy(_));
            let source_ranges = match &entry.central {
                PreparedCentral::Promoted(promotions) => promotions
                    .first()
                    .and_then(|promotion| promotion.source_ranges.as_deref()),
                _ => None,
            };
            if let Err(source) = write_prepared_central(
                &mut sink,
                central,
                patch,
                is_unchanged,
                source_ranges,
                accounting,
            ) {
                return Err(map_emit_error(source, &mut sink));
            }
        }
        if let Err(source) = write_prepared_tail(&layout.tail, &mut sink, accounting) {
            return Err(map_emit_error(source, &mut sink));
        }
        if let Err(source) = crate::accounting::write_all_counted(
            &mut sink,
            &self.archive_comment,
            accounting,
            AccountingWriteKind::RawUnchangedSource,
        ) {
            return Err(map_emit_error(source, &mut sink));
        }
        if sink.accepted != layout.output_size {
            return Err(ReplayPublicationError::Archive {
                pass: ReplayPass::Emit,
                source: unsupported("replay output size did not match preflight"),
                callback_error: None,
                progress: sink.progress(),
            });
        }
        if let Err(source) = sink.flush() {
            return Err(sink_error(ReplayPass::Emit, source, None, &mut sink));
        }
        if let Some(source) = sink.take_error() {
            return Err(ReplayPublicationError::Sink {
                pass: ReplayPass::Emit,
                source,
                callback_error: None,
                progress: sink.progress(),
            });
        }
        Ok(sink.into_inner())
    }
}

fn replay_target_name(entry: &super::PreservedEntry) -> Result<&str, Error> {
    if entry.local_central_name_mismatch {
        return Err(unsupported("replay target local and central names differ"));
    }
    let raw_name = entry.raw_name_bytes();
    let name = std::str::from_utf8(raw_name)
        .map_err(|_| unsupported("replay target name is not UTF-8"))?;
    let normalized = ZipFilePath::from_str(name);
    if normalized.is_dir() || normalized.as_ref().as_bytes() != raw_name {
        return Err(unsupported("replay target name is not already normalized"));
    }
    Ok(name)
}

fn prepare_replay_target(
    name: &str,
    compression: CompressionMethod,
    measurement: ReplayMeasurement,
) -> Result<PreparedEntry, Error> {
    let mut member = prepare_sized_member(
        name,
        compression,
        measurement.decoded_crc32,
        measurement.compressed_size,
        measurement.decoded_size,
        0,
    )?;
    let framing = member.local_bytes()?;
    let central = member.central_bytes()?;
    let prepared = PreparedEntry {
        local: PreparedLocal::Replay {
            framing,
            compressed_len: measurement.compressed_size,
        },
        central: PreparedCentral::Generated(central),
        generated_payload: None,
        omitted: false,
    };
    Ok(prepared)
}

fn measure_callback<F, E>(
    callback: &mut F,
    compression: CompressionMethod,
    limits: ReplayLimits,
) -> Result<ReplayMeasurement, ReplayPublicationError<E>>
where
    F: FnMut(&mut dyn Write) -> Result<(), E>,
{
    let mut output = ReplayPayloadWriter::new(compression, limits, ReplayMeter);
    let callback_result = callback(&mut output);
    if let Err(source) = callback_result {
        return Err(output.callback_error(ReplayPass::Measure, source));
    }
    match output.finish() {
        Ok(measurement) => Ok(measurement),
        Err(failure) => Err(map_writer_failure(
            ReplayPass::Measure,
            failure,
            None,
            ReplayProgress::Untouched,
        )),
    }
}

fn emit_callback<F, E, W>(
    callback: &mut F,
    compression: CompressionMethod,
    limits: ReplayLimits,
    sink: &mut ReplaySink<W>,
    accounting: &mut ZipOperationAccounting,
) -> Result<ReplayMeasurement, ReplayPublicationError<E>>
where
    F: FnMut(&mut dyn Write) -> Result<(), E>,
    W: Write,
{
    let accounting_kind = match compression {
        CompressionMethod::Store => AccountingWriteKind::Stored,
        CompressionMethod::Deflate => AccountingWriteKind::GeneratedDeflate,
        _ => {
            return Err(ReplayPublicationError::Archive {
                pass: ReplayPass::Emit,
                source: ErrorKind::UnsupportedCompressionMethod(compression.as_id().as_u16())
                    .into(),
                callback_error: None,
                progress: sink.progress(),
            });
        },
    };
    let mut output = ReplayPayloadWriter::with_accounting(
        compression,
        limits,
        &mut *sink,
        accounting,
        accounting_kind,
    );
    let callback_result = callback(&mut output);
    if let Err(source) = callback_result {
        let failure = output.failure.take();
        drop(output);
        if let Some(sink_error) = sink.take_error() {
            return Err(ReplayPublicationError::Sink {
                pass: ReplayPass::Emit,
                source: sink_error,
                callback_error: Some(source),
                progress: sink.progress(),
            });
        }
        if let Some(failure) = failure {
            return Err(map_writer_failure(
                ReplayPass::Emit,
                failure,
                Some(source),
                sink.progress(),
            ));
        }
        return Err(ReplayPublicationError::Callback {
            pass: ReplayPass::Emit,
            source,
            progress: sink.progress(),
        });
    }
    let result = output.finish();
    drop(output);
    match result {
        Ok(measurement) => Ok(measurement),
        Err(failure) => Err(if let Some(source) = sink.take_error() {
            ReplayPublicationError::Sink {
                pass: ReplayPass::Emit,
                source,
                callback_error: None,
                progress: sink.progress(),
            }
        } else {
            map_writer_failure(ReplayPass::Emit, failure, None, sink.progress())
        }),
    }
}

#[derive(Debug)]
enum ReplayWriterFailure {
    Sink(io::Error),
    Limit {
        resource: ReplayResource,
        actual: u64,
        maximum: u64,
    },
    Compression(io::Error),
}

struct ReplayPayloadWriter<'sink, W> {
    compression: CompressionMethod,
    limits: ReplayLimits,
    sink: W,
    accounting: Option<(&'sink mut ZipOperationAccounting, AccountingWriteKind)>,
    compressor: Option<Compress>,
    input_window: [u8; REPLAY_INPUT_WINDOW_SIZE],
    input_window_len: usize,
    decoded_size: u64,
    compressed_size: u64,
    decoded_crc32: u32,
    decoded_sha256: Sha256,
    compressed_sha256: Sha256,
    failure: Option<ReplayWriterFailure>,
    finished: bool,
}

impl<'sink, W> ReplayPayloadWriter<'sink, W>
where
    W: Write,
{
    fn new(compression: CompressionMethod, limits: ReplayLimits, sink: W) -> Self {
        Self::with_optional_accounting(compression, limits, sink, None)
    }

    fn with_accounting(
        compression: CompressionMethod,
        limits: ReplayLimits,
        sink: W,
        accounting: &'sink mut ZipOperationAccounting,
        kind: AccountingWriteKind,
    ) -> Self {
        Self::with_optional_accounting(compression, limits, sink, Some((accounting, kind)))
    }

    fn with_optional_accounting(
        compression: CompressionMethod,
        limits: ReplayLimits,
        sink: W,
        accounting: Option<(&'sink mut ZipOperationAccounting, AccountingWriteKind)>,
    ) -> Self {
        Self {
            compression,
            limits,
            sink,
            accounting,
            compressor: (compression == CompressionMethod::Deflate)
                .then(|| Compress::new(Compression::default(), false)),
            input_window: [0; REPLAY_INPUT_WINDOW_SIZE],
            input_window_len: 0,
            decoded_size: 0,
            compressed_size: 0,
            decoded_crc32: 0,
            decoded_sha256: Sha256::new(),
            compressed_sha256: Sha256::new(),
            failure: None,
            finished: false,
        }
    }

    fn callback_error<E>(&mut self, pass: ReplayPass, source: E) -> ReplayPublicationError<E> {
        if let Some(failure) = self.failure.take() {
            return map_writer_failure(pass, failure, Some(source), ReplayProgress::Untouched);
        }
        ReplayPublicationError::Callback {
            pass,
            source,
            progress: ReplayProgress::Untouched,
        }
    }

    fn remember_failure(&mut self, failure: ReplayWriterFailure) {
        if self.failure.is_none() {
            self.failure = Some(failure);
        }
    }

    fn fail_with(&mut self, failure: ReplayWriterFailure) -> io::Error {
        let returned = failure_io(&failure);
        self.remember_failure(failure);
        returned
    }

    fn fail_compression(&mut self, error: io::Error) -> io::Error {
        let returned = clone_io_error(Some(&error)).unwrap_or_else(|| {
            io::Error::new(io::ErrorKind::InvalidData, "replay compression failed")
        });
        self.remember_failure(ReplayWriterFailure::Compression(error));
        returned
    }

    fn finish(&mut self) -> Result<ReplayMeasurement, ReplayWriterFailure> {
        if self.finished {
            return Err(ReplayWriterFailure::Compression(io::Error::new(
                io::ErrorKind::InvalidInput,
                "replay output finished twice",
            )));
        }
        self.finished = true;
        if let Some(failure) = self.failure.take() {
            return Err(failure);
        }
        if self.compression == CompressionMethod::Deflate {
            if self.input_window_len != 0 {
                let input_len = self.input_window_len;
                if let Err(error) = self.compress_input_window(input_len) {
                    return Err(match self.failure.take() {
                        Some(failure) => failure,
                        None => ReplayWriterFailure::Compression(error),
                    });
                }
                self.input_window_len = 0;
            }
            let Some(mut compressor) = self.compressor.take() else {
                return Err(ReplayWriterFailure::Compression(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "replay Deflate compressor is unavailable",
                )));
            };
            let mut output = [0u8; REPLAY_COMPRESS_BUFFER_SIZE];
            loop {
                let before = compressor.total_out();
                let status = match compressor.compress(&[], &mut output, FlushCompress::Finish) {
                    Ok(status) => status,
                    Err(error) => {
                        return Err(ReplayWriterFailure::Compression(io::Error::new(
                            io::ErrorKind::InvalidData,
                            error.to_string(),
                        )));
                    },
                };
                let produced = checked_delta(compressor.total_out(), before)?;
                if produced > output.len() {
                    return Err(ReplayWriterFailure::Compression(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "replay compressor produced beyond its output buffer",
                    )));
                }
                if produced != 0 {
                    if let Err(error) = self.write_compressed(&output[..produced]) {
                        return Err(match self.failure.take() {
                            Some(failure) => failure,
                            None => ReplayWriterFailure::Compression(error),
                        });
                    }
                }
                if status == Status::StreamEnd {
                    break;
                }
                if produced == 0 {
                    return Err(ReplayWriterFailure::Compression(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "Deflate replay did not make finish progress",
                    )));
                }
            }
        }
        if let Some(failure) = self.failure.take() {
            return Err(failure);
        }
        Ok(self.measurement())
    }

    fn measurement(&self) -> ReplayMeasurement {
        ReplayMeasurement {
            decoded_size: self.decoded_size,
            compressed_size: self.compressed_size,
            decoded_crc32: self.decoded_crc32,
            decoded_sha256: self.decoded_sha256.clone().finalize().into(),
            compressed_sha256: self.compressed_sha256.clone().finalize().into(),
        }
    }

    fn write_stored(&mut self, input: &[u8]) -> io::Result<usize> {
        self.ensure_decoded_limit(input.len())?;
        self.ensure_compressed_limit(input.len())?;
        let mut accepted = 0usize;
        match write_all_replay(&mut self.sink, input, &mut self.failure, &mut accepted) {
            Ok(_) => {
                self.record_decoded(input)?;
                self.record_compressed(input)?;
                Ok(input.len())
            },
            Err(error) => {
                if accepted != 0 {
                    let prefix = &input[..accepted];
                    let _ = self.record_decoded(prefix);
                    let _ = self.record_compressed(prefix);
                }
                Err(error)
            },
        }
    }

    fn write_deflate(&mut self, input: &[u8]) -> io::Result<usize> {
        self.ensure_decoded_limit(input.len())?;
        self.record_decoded(input)?;
        let mut offset = 0usize;
        while offset < input.len() {
            let available = REPLAY_INPUT_WINDOW_SIZE - self.input_window_len;
            let copied = available.min(input.len() - offset);
            let end = self.input_window_len + copied;
            self.input_window[self.input_window_len..end]
                .copy_from_slice(&input[offset..offset + copied]);
            self.input_window_len = end;
            offset += copied;
            if self.input_window_len == REPLAY_INPUT_WINDOW_SIZE {
                let input_len = self.input_window_len;
                self.compress_input_window(input_len)?;
                self.input_window_len = 0;
            }
        }
        Ok(input.len())
    }

    fn compress_input_window(&mut self, input_len: usize) -> io::Result<()> {
        if input_len == 0 {
            return Ok(());
        }
        if input_len != self.input_window_len || input_len > REPLAY_INPUT_WINDOW_SIZE {
            return Err(self.fail_compression(io::Error::new(
                io::ErrorKind::InvalidInput,
                "invalid replay Deflate input window",
            )));
        }
        let mut consumed = 0usize;
        let mut output = [0u8; REPLAY_COMPRESS_BUFFER_SIZE];
        while consumed < input_len {
            let (status, before_in, before_out, after_in, after_out) = {
                let Some(compressor) = self.compressor.as_mut() else {
                    return Err(self.fail_compression(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "replay Deflate compressor is unavailable",
                    )));
                };
                let before_in = compressor.total_in();
                let before_out = compressor.total_out();
                let status = compressor
                    .compress(
                        &self.input_window[consumed..input_len],
                        &mut output,
                        FlushCompress::None,
                    )
                    .map_err(|error| io::Error::new(io::ErrorKind::InvalidData, error.to_string()));
                let after_in = compressor.total_in();
                let after_out = compressor.total_out();
                (status, before_in, before_out, after_in, after_out)
            };
            let status = match status {
                Ok(status) => status,
                Err(error) => return Err(self.fail_compression(error)),
            };
            let consumed_now = match checked_delta(after_in, before_in) {
                Ok(delta) => delta,
                Err(failure) => return Err(self.fail_with(failure)),
            };
            let produced = match checked_delta(after_out, before_out) {
                Ok(delta) => delta,
                Err(failure) => return Err(self.fail_with(failure)),
            };
            if produced > output.len() {
                return Err(
                    self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "replay compressor produced beyond its output buffer",
                    ))),
                );
            }
            if produced != 0 {
                self.write_compressed(&output[..produced])?;
            }
            let remaining = input_len - consumed;
            if consumed_now > remaining {
                return Err(
                    self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "replay compressor consumed beyond its input window",
                    ))),
                );
            }
            consumed += consumed_now;
            if consumed_now == 0 && produced == 0 {
                return Err(
                    self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "Deflate replay did not make input progress",
                    ))),
                );
            }
            if status == Status::StreamEnd {
                return Err(
                    self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "Deflate replay ended before all input was consumed",
                    ))),
                );
            }
        }
        Ok(())
    }

    fn write_compressed(&mut self, output: &[u8]) -> io::Result<()> {
        self.ensure_compressed_limit(output.len())?;
        let mut accepted = 0usize;
        match write_all_replay(&mut self.sink, output, &mut self.failure, &mut accepted) {
            Ok(_) => self.record_compressed(output),
            Err(error) => {
                if accepted != 0 {
                    let _ = self.record_compressed(&output[..accepted]);
                }
                Err(error)
            },
        }
    }

    fn ensure_decoded_limit(&mut self, input_len: usize) -> io::Result<()> {
        let requested = match u64::try_from(input_len) {
            Ok(requested) => requested,
            Err(_) => {
                return Err(
                    self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "replay decoded size overflow",
                    ))),
                );
            },
        };
        let Some(next) = self.decoded_size.checked_add(requested) else {
            return Err(
                self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "replay decoded size overflow",
                ))),
            );
        };
        if next > self.limits.max_decoded_bytes {
            return Err(self.fail_with(ReplayWriterFailure::Limit {
                resource: ReplayResource::DecodedBytes,
                actual: next,
                maximum: self.limits.max_decoded_bytes,
            }));
        }
        Ok(())
    }

    fn ensure_compressed_limit(&mut self, output_len: usize) -> io::Result<()> {
        let requested = match u64::try_from(output_len) {
            Ok(requested) => requested,
            Err(_) => {
                return Err(
                    self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "replay compressed size overflow",
                    ))),
                );
            },
        };
        let Some(next) = self.compressed_size.checked_add(requested) else {
            return Err(
                self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "replay compressed size overflow",
                ))),
            );
        };
        if next > self.limits.max_compressed_bytes {
            return Err(self.fail_with(ReplayWriterFailure::Limit {
                resource: ReplayResource::CompressedBytes,
                actual: next,
                maximum: self.limits.max_compressed_bytes,
            }));
        }
        Ok(())
    }

    fn record_decoded(&mut self, bytes: &[u8]) -> io::Result<()> {
        let length = match u64::try_from(bytes.len()) {
            Ok(length) => length,
            Err(_) => {
                return Err(
                    self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "replay decoded size overflow",
                    ))),
                );
            },
        };
        let next = match self.decoded_size.checked_add(length) {
            Some(next) => next,
            None => {
                return Err(
                    self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "replay decoded size overflow",
                    ))),
                );
            },
        };
        if next > self.limits.max_decoded_bytes {
            return Err(self.fail_with(ReplayWriterFailure::Limit {
                resource: ReplayResource::DecodedBytes,
                actual: next,
                maximum: self.limits.max_decoded_bytes,
            }));
        }
        self.decoded_size = next;
        self.decoded_crc32 = crc32_chunk(bytes, self.decoded_crc32);
        self.decoded_sha256.update(bytes);
        Ok(())
    }

    fn record_compressed(&mut self, bytes: &[u8]) -> io::Result<()> {
        let length = match u64::try_from(bytes.len()) {
            Ok(length) => length,
            Err(_) => {
                return Err(
                    self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "replay compressed size overflow",
                    ))),
                );
            },
        };
        let next = match self.compressed_size.checked_add(length) {
            Some(next) => next,
            None => {
                return Err(
                    self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "replay compressed size overflow",
                    ))),
                );
            },
        };
        if next > self.limits.max_compressed_bytes {
            return Err(self.fail_with(ReplayWriterFailure::Limit {
                resource: ReplayResource::CompressedBytes,
                actual: next,
                maximum: self.limits.max_compressed_bytes,
            }));
        }
        self.compressed_size = next;
        self.compressed_sha256.update(bytes);
        let accounting_error = if let Some((accounting, kind)) = &mut self.accounting {
            match u64::try_from(bytes.len()) {
                Ok(count) => (*kind).add(accounting, count).err(),
                Err(_) => Some(Error::from(ErrorKind::InvalidInput {
                    msg: "replay accounting byte count overflow".to_string(),
                })),
            }
        } else {
            None
        };
        if let Some(error) = accounting_error {
            return Err(
                self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                    io::ErrorKind::InvalidData,
                    error.to_string(),
                ))),
            );
        }
        Ok(())
    }
}

impl<W: Write> Write for ReplayPayloadWriter<'_, W> {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if self.finished {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "replay output was already finalized",
            ));
        }
        if self.failure.is_some() {
            return Err(io::Error::new(
                io::ErrorKind::Other,
                "replay output has failed",
            ));
        }
        if input.is_empty() {
            return Ok(0);
        }
        match self.compression {
            CompressionMethod::Store => self.write_stored(input),
            CompressionMethod::Deflate => self.write_deflate(input),
            _ => Err(
                self.fail_with(ReplayWriterFailure::Compression(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "unsupported replay compression method",
                ))),
            ),
        }
    }

    fn flush(&mut self) -> io::Result<()> {
        if let Some(failure) = self.failure.as_ref() {
            return Err(failure_io(failure));
        }
        if let Err(error) = self.sink.flush() {
            let returned = clone_io_error(Some(&error)).unwrap_or_else(|| {
                io::Error::new(io::ErrorKind::Other, "replay sink flush failed")
            });
            self.remember_failure(ReplayWriterFailure::Sink(error));
            return Err(returned);
        }
        Ok(())
    }
}

struct ReplayMeter;

impl Write for ReplayMeter {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        Ok(buffer.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct ReplaySink<W> {
    inner: W,
    accepted: u64,
    expected: Option<u64>,
    overreported: bool,
    flush_started: bool,
    flush_completed: bool,
    first_error: Option<io::Error>,
}

impl<W> ReplaySink<W> {
    fn new(inner: W) -> Self {
        Self {
            inner,
            accepted: 0,
            expected: None,
            overreported: false,
            flush_started: false,
            flush_completed: false,
            first_error: None,
        }
    }

    fn remember_error(&mut self, error: io::Error) {
        if self.first_error.is_none() {
            self.first_error = Some(error);
        }
    }

    fn set_expected(&mut self, expected: u64) {
        self.expected = Some(expected);
    }

    fn progress(&self) -> ReplayProgress {
        if self.overreported {
            ReplayProgress::Indeterminate {
                accepted_before: self.accepted,
            }
        } else if self.accepted == 0 {
            ReplayProgress::Untouched
        } else if self.expected == Some(self.accepted) {
            if self.flush_completed {
                ReplayProgress::Complete {
                    bytes: self.accepted,
                }
            } else if self.flush_started {
                ReplayProgress::CompleteUnflushed {
                    bytes: self.accepted,
                }
            } else {
                ReplayProgress::Prefix {
                    accepted: self.accepted,
                }
            }
        } else {
            ReplayProgress::Prefix {
                accepted: self.accepted,
            }
        }
    }

    fn take_error(&mut self) -> Option<io::Error> {
        self.first_error.take()
    }

    fn into_inner(self) -> W {
        self.inner
    }
}

impl<W: Write> Write for ReplaySink<W> {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        if !buffer.is_empty() {
            // A successful flush only covers bytes accepted before this write.
            // Clear both markers before touching the destination so an error
            // after an early callback flush cannot be reported as complete.
            self.flush_started = false;
            self.flush_completed = false;
        }
        let written = match self.inner.write(buffer) {
            Ok(written) => written,
            Err(error) if error.kind() == io::ErrorKind::Interrupted => return Err(error),
            Err(error) => {
                let kind = error.kind();
                let message = error.to_string();
                self.remember_error(error);
                return Err(io::Error::new(kind, message));
            },
        };
        if written > buffer.len() {
            self.overreported = true;
            let error = io::Error::new(
                io::ErrorKind::InvalidData,
                "replay sink reported more bytes than supplied",
            );
            self.remember_error(io::Error::new(error.kind(), error.to_string()));
            return Err(error);
        }
        let written_u64 = match u64::try_from(written) {
            Ok(written) => written,
            Err(_) => {
                let error = io::Error::new(
                    io::ErrorKind::InvalidData,
                    "replay output count overflows u64",
                );
                self.remember_error(io::Error::new(error.kind(), error.to_string()));
                return Err(error);
            },
        };
        self.accepted = match self.accepted.checked_add(written_u64) {
            Some(accepted) => accepted,
            None => {
                let error = io::Error::new(
                    io::ErrorKind::InvalidData,
                    "replay output count overflows u64",
                );
                self.remember_error(io::Error::new(error.kind(), error.to_string()));
                return Err(error);
            },
        };
        if written == 0 && !buffer.is_empty() {
            let error = io::Error::new(io::ErrorKind::WriteZero, "replay sink accepted no bytes");
            self.remember_error(io::Error::new(error.kind(), error.to_string()));
            return Err(error);
        }
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.flush_started = true;
        match self.inner.flush() {
            Ok(()) => {
                self.flush_completed = true;
                Ok(())
            },
            Err(error) => {
                self.flush_completed = false;
                let kind = error.kind();
                let message = error.to_string();
                self.remember_error(error);
                Err(io::Error::new(kind, message))
            },
        }
    }
}

#[cfg(test)]
pub(super) fn test_early_flush_then_late_write_failure_progress() -> ReplayProgress {
    struct FailOnSecondWrite {
        writes: usize,
    }

    impl Write for FailOnSecondWrite {
        fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
            self.writes += 1;
            if self.writes == 2 {
                return Err(io::Error::new(
                    io::ErrorKind::BrokenPipe,
                    "late replay sink failure",
                ));
            }
            Ok(buffer.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    let mut sink = ReplaySink::new(FailOnSecondWrite { writes: 0 });
    sink.set_expected(1);
    {
        let mut output =
            ReplayPayloadWriter::new(CompressionMethod::Store, ReplayLimits::default(), &mut sink);
        output
            .write_all(&[1])
            .expect("the first replay byte should be accepted");
        output
            .flush()
            .expect("the early replay flush should succeed");
        assert!(output.write_all(&[2]).is_err());
    }
    sink.progress()
}

fn write_all_replay<W: Write>(
    sink: &mut W,
    mut bytes: &[u8],
    failure: &mut Option<ReplayWriterFailure>,
    accepted: &mut usize,
) -> io::Result<usize> {
    let mut written_total = 0usize;
    while !bytes.is_empty() {
        let written = match sink.write(bytes) {
            Ok(written) => written,
            Err(error) if error.kind() == io::ErrorKind::Interrupted => continue,
            Err(error) => {
                let kind = error.kind();
                let message = error.to_string();
                if failure.is_none() {
                    *failure = Some(ReplayWriterFailure::Sink(error));
                }
                return Err(io::Error::new(kind, message));
            },
        };
        if written == 0 {
            let error = io::Error::new(io::ErrorKind::WriteZero, "replay sink accepted no bytes");
            if failure.is_none() {
                *failure = Some(ReplayWriterFailure::Sink(io::Error::new(
                    error.kind(),
                    error.to_string(),
                )));
            }
            return Err(error);
        }
        if written > bytes.len() {
            let error = io::Error::new(
                io::ErrorKind::InvalidData,
                "replay sink reported more bytes than supplied",
            );
            if failure.is_none() {
                *failure = Some(ReplayWriterFailure::Sink(io::Error::new(
                    error.kind(),
                    error.to_string(),
                )));
            }
            return Err(error);
        }
        written_total = match written_total.checked_add(written) {
            Some(total) => total,
            None => {
                let error = io::Error::new(
                    io::ErrorKind::InvalidData,
                    "replay output count overflows usize",
                );
                if failure.is_none() {
                    *failure = Some(ReplayWriterFailure::Sink(io::Error::new(
                        error.kind(),
                        error.to_string(),
                    )));
                }
                return Err(error);
            },
        };
        *accepted = written_total;
        bytes = &bytes[written..];
    }
    Ok(written_total)
}

fn checked_delta(current: u64, previous: u64) -> Result<usize, ReplayWriterFailure> {
    let delta = current.checked_sub(previous).ok_or_else(|| {
        ReplayWriterFailure::Compression(io::Error::new(
            io::ErrorKind::InvalidData,
            "replay compressor counter moved backwards",
        ))
    })?;
    usize::try_from(delta).map_err(|_| {
        ReplayWriterFailure::Compression(io::Error::new(
            io::ErrorKind::InvalidData,
            "replay compressor output does not fit usize",
        ))
    })
}

fn map_writer_failure<E>(
    pass: ReplayPass,
    failure: ReplayWriterFailure,
    callback_error: Option<E>,
    progress: ReplayProgress,
) -> ReplayPublicationError<E> {
    match failure {
        ReplayWriterFailure::Sink(source) => ReplayPublicationError::Sink {
            pass,
            source,
            callback_error,
            progress,
        },
        ReplayWriterFailure::Limit {
            resource,
            actual,
            maximum,
        } => ReplayPublicationError::Limit {
            pass,
            resource,
            actual,
            maximum,
            callback_error,
            progress,
        },
        ReplayWriterFailure::Compression(source) => ReplayPublicationError::Archive {
            pass,
            source: source.into(),
            callback_error,
            progress,
        },
    }
}

fn failure_io(failure: &ReplayWriterFailure) -> io::Error {
    match failure {
        ReplayWriterFailure::Sink(source) | ReplayWriterFailure::Compression(source) => {
            clone_io_error(Some(source))
                .unwrap_or_else(|| io::Error::new(io::ErrorKind::Other, "replay writer failed"))
        },
        ReplayWriterFailure::Limit { resource, .. } => io::Error::new(
            io::ErrorKind::InvalidInput,
            format!("replay {resource} limit exceeded"),
        ),
    }
}

fn sink_error<E, W: Write>(
    pass: ReplayPass,
    fallback: io::Error,
    callback_error: Option<E>,
    sink: &mut ReplaySink<W>,
) -> ReplayPublicationError<E> {
    ReplayPublicationError::Sink {
        pass,
        source: sink.take_error().unwrap_or(fallback),
        callback_error,
        progress: sink.progress(),
    }
}

fn map_emit_error<E, W: Write>(
    source: Error,
    sink: &mut ReplaySink<W>,
) -> ReplayPublicationError<E> {
    if let Some(sink_error) = sink.take_error() {
        ReplayPublicationError::Sink {
            pass: ReplayPass::Emit,
            source: sink_error,
            callback_error: None,
            progress: sink.progress(),
        }
    } else {
        ReplayPublicationError::Archive {
            pass: ReplayPass::Emit,
            source,
            callback_error: None,
            progress: sink.progress(),
        }
    }
}

fn clone_io_error(source: Option<&io::Error>) -> Option<io::Error> {
    source.map(|error| io::Error::new(error.kind(), error.to_string()))
}

fn unsupported(reason: &'static str) -> Error {
    ErrorKind::UnsupportedPreservation { reason }.into()
}
