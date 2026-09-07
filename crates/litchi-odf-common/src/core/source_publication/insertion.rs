//! Bounded source-backed insertion into the decoded `content.xml` member.
//!
//! The insertion plan is intentionally separate from the ordinary raw
//! replacement publisher.  It retains only the authored fragment and hashes
//! of the source/candidate content; both source and candidate bytes are
//! streamed again for ZIP replay.  No complete XML member is retained.

use super::super::package::{SourceBackedPackage, SourceMemberReaderError};
use super::super::private::{BindingTracker, XmlStreamEvent, XmlStreamLimits};
use super::super::xml_splice::AuthoredXmlFragment;
use super::{
    CheckedSink, OutputState, SOURCE_CONTENT_OPERATION_BUFFER_BYTES,
    SOURCE_CONTENT_SCAN_CANCELLED_MESSAGE, SourceContentPublicationError,
    SourceContentPublicationOptions, SourceContentPublicationProgress,
    SourceContentPublicationReport, check_cancellation, consume_execution_resource,
    preflight_insertion_source, reconcile_source_state, reserve_memory, scan_content_reader,
};
use crate::constants;
use litchi_core::{Error, ExecutionError, Reservation, Resource, Result, SourceVersion};
use sha2::{Digest, Sha256};
use soapberry_zip::{CompressionMethod, ReplayLimits, ReplayPublicationError};
use std::borrow::Borrow;
use std::fmt;
use std::io::{self, BufRead, Read, Write};
use std::sync::Arc;

const COPY_CHUNK_SIZE: usize = 64 * 1024;
const INSERTION_OBJECT_BYTES: u64 = 512;

/// The callback failure carried through a ZIP replay.
///
/// The type is public so callers can inspect the source-member error retained
/// by an abortable replay callback without relying on formatted text.  It is
/// hidden from ordinary documentation because it is an advanced publication
/// diagnostic rather than a CRUD value.
#[doc(hidden)]
#[derive(Debug)]
#[non_exhaustive]
pub enum SourceContentInsertionCallbackError {
    /// A verified source-member read failed.  The box permits the callback's
    /// own typed error to remain available in the nested reader diagnostic.
    Member(Box<SourceMemberReaderError<Self>>),
    /// A source or destination `Read`/`Write` operation failed.
    Io(io::Error),
    /// The operation was cancelled between bounded chunks.
    Cancelled,
    /// The execution context rejected the operation.
    Execution(ExecutionError),
    /// A replay pass observed bytes different from the prepared plan.
    Verification {
        /// Which hash/length proof failed.
        what: &'static str,
    },
}

impl fmt::Display for SourceContentInsertionCallbackError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Member(error) => error.fmt(formatter),
            Self::Io(error) => error.fmt(formatter),
            Self::Cancelled => formatter.write_str("source-content insertion cancelled"),
            Self::Execution(error) => error.fmt(formatter),
            Self::Verification { what } => {
                write!(
                    formatter,
                    "source-content insertion replay verification failed: {what}"
                )
            },
        }
    }
}

impl std::error::Error for SourceContentInsertionCallbackError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Member(error) => Some(error.as_ref()),
            Self::Io(error) => Some(error),
            Self::Execution(error) => Some(error),
            Self::Cancelled | Self::Verification { .. } => None,
        }
    }
}

/// A typed failure from bounded source-content insertion publication.
///
/// Replay failures retain the ZIP replay error and the exact sink progress
/// observed by this layer.  Preflight failures retain the ordinary typed
/// publication error.  The two error families are deliberately not flattened
/// into strings because a replay callback may contain a source-member error
/// and a secondary callback error at the same time.
#[derive(Debug)]
#[non_exhaustive]
pub enum SourceContentInsertionError {
    /// Common ODF/source preflight or post-publication failure.
    Publication {
        /// Exact bytes accepted by the caller sink before failure.
        progress: SourceContentPublicationProgress,
        /// The typed common publication cause.
        source: SourceContentPublicationError,
    },
    /// ZIP replay or replay-callback failure.
    Replay {
        /// Exact bytes accepted by the caller sink before failure.
        progress: SourceContentPublicationProgress,
        /// The typed ZIP replay cause, including callback diagnostics.
        source: ReplayPublicationError<SourceContentInsertionCallbackError>,
    },
    /// A ZIP replay transport failure whose public ODF publication cause was
    /// recognized while retaining the complete typed replay diagnostic,
    /// including any secondary callback failure.
    Transport {
        /// Exact bytes accepted by the caller sink before failure.
        progress: SourceContentPublicationProgress,
        /// The public cancellation, limit, source, execution, or sink cause.
        publication: SourceContentPublicationError,
        /// The original replay cause, including callback diagnostics.
        replay: Box<ReplayPublicationError<SourceContentInsertionCallbackError>>,
    },
}

impl SourceContentInsertionError {
    /// Return the exact sink progress at failure.
    #[must_use]
    pub const fn progress(&self) -> SourceContentPublicationProgress {
        match self {
            Self::Publication { progress, .. }
            | Self::Replay { progress, .. }
            | Self::Transport { progress, .. } => *progress,
        }
    }

    /// Return bytes definitely accepted by the caller sink before failure.
    #[must_use]
    pub const fn written(&self) -> u64 {
        self.progress().accepted()
    }
}

impl fmt::Display for SourceContentInsertionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Publication { source, progress } => {
                write!(
                    formatter,
                    "source-content insertion publication failed ({progress:?}): {source}"
                )
            },
            Self::Replay { source, progress } => {
                write!(
                    formatter,
                    "source-content insertion ZIP replay failed ({progress:?}): {source}"
                )
            },
            Self::Transport {
                publication,
                replay,
                progress,
            } => write!(
                formatter,
                "source-content insertion transport failed ({progress:?}): {publication}; replay: {replay}"
            ),
        }
    }
}

impl std::error::Error for SourceContentInsertionError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Publication { source, .. } => Some(source),
            Self::Replay { source, .. } => Some(source),
            Self::Transport { publication, .. } => Some(publication),
        }
    }
}

impl From<SourceContentPublicationError> for SourceContentInsertionError {
    fn from(source: SourceContentPublicationError) -> Self {
        let progress = source.progress();
        Self::Publication { progress, source }
    }
}

/// A prepared bounded insertion at a decoded-byte offset in `content.xml`.
///
/// The plan retains the source snapshot and authored fragment, source/target
/// SHA-256 proofs, and lengths.  It never retains complete source or target
/// XML bytes.
#[derive(Debug)]
pub struct SourceContentInsertionPlan {
    source: Arc<SourceBackedPackage>,
    source_version: SourceVersion,
    source_length: u64,
    decoded_offset: u64,
    fragment: AuthoredXmlFragment,
    _retained_memory: Option<Reservation>,
    source_content_sha256: [u8; 32],
    target_content_sha256: [u8; 32],
    source_content_length: u64,
    target_content_length: u64,
}

impl SourceContentInsertionPlan {
    /// Prepare an insertion after streaming and validating the candidate XML.
    pub fn prepare(
        source: Arc<SourceBackedPackage>,
        decoded_offset: u64,
        fragment: AuthoredXmlFragment,
        limits: XmlStreamLimits,
        options: &SourceContentPublicationOptions,
    ) -> std::result::Result<Self, SourceContentPublicationError> {
        Self::prepare_with_validator(
            source,
            decoded_offset,
            fragment,
            limits,
            options,
            |_event, _bindings| Ok(()),
        )
    }

    /// Prepare an insertion while visiting every candidate XML event.
    ///
    /// The visitor runs during the same bounded candidate scan that proves
    /// the whole XML document.  It is suitable for a format owner such as
    /// ODP to prove page-tail semantics without materializing the candidate.
    pub fn prepare_with_validator<F>(
        source: Arc<SourceBackedPackage>,
        decoded_offset: u64,
        input_fragment: AuthoredXmlFragment,
        limits: XmlStreamLimits,
        options: &SourceContentPublicationOptions,
        mut validator: F,
    ) -> std::result::Result<Self, SourceContentPublicationError>
    where
        F: for<'event> FnMut(&'event XmlStreamEvent<'event>, &BindingTracker) -> Result<()>,
    {
        check_cancellation(options, SourceContentPublicationProgress::Untouched)?;
        let fragment_len = u64::try_from(input_fragment.bytes().len()).map_err(|_| {
            SourceContentPublicationError::Unsupported {
                reason: "authored insertion fragment length exceeds u64".to_string(),
            }
        })?;
        validate_replacement_limit(fragment_len, options)?;
        let retained_fragment_bytes = retained_fragment_bytes(&input_fragment)?;
        let retained_memory = reserve_memory(
            options,
            retained_fragment_bytes,
            SourceContentPublicationProgress::Untouched,
        )?;
        // Move the fragment into a local declared after the lease so every
        // early-return path drops its backing Vec before releasing the charge.
        let fragment = input_fragment;

        source.ensure_current_for_publication().map_err(|error| {
            super::map_core_error(error, SourceContentPublicationProgress::Untouched)
        })?;
        super::reject_encrypted_source(&source)?;

        let source_content_length = source
            .publication_member_size(constants::ODF_CONTENT)
            .map_err(|error| {
                super::map_core_error(error, SourceContentPublicationProgress::Untouched)
            })?;
        if decoded_offset > source_content_length {
            return Err(SourceContentPublicationError::Core(Error::InvalidFormat(
                "content.xml insertion offset exceeds decoded source length".to_string(),
            )));
        }
        let target_content_length =
            source_content_length
                .checked_add(fragment_len)
                .ok_or_else(|| SourceContentPublicationError::Unsupported {
                    reason: "content.xml insertion target length overflows u64".to_string(),
                })?;
        validate_replacement_limit(target_content_length, options)?;

        let entry_count = u64::try_from(source.publication_entry_count()).unwrap_or(u64::MAX);
        let xml_memory = limits.memory_upper_bound().map_err(|error| {
            super::map_core_error(error, SourceContentPublicationProgress::Untouched)
        })?;
        let memory = insertion_transient_memory_requirement(
            source.publication_metadata_bytes(),
            entry_count,
            xml_memory,
        )?;
        let _transient_memory =
            reserve_memory(options, memory, SourceContentPublicationProgress::Untouched)?;
        let _input = source
            .begin_publication_input_accounting(options.execution_context())
            .map_err(|source| SourceContentPublicationError::Allocation {
                resource: "source-content insertion input-accounting stack",
                source,
            })?;
        let object_work = entry_count.checked_add(2).ok_or_else(|| {
            SourceContentPublicationError::Unsupported {
                reason: "source-content insertion object accounting overflow".to_string(),
            }
        })?;
        consume_execution_resource(
            options,
            Resource::Objects,
            object_work,
            SourceContentPublicationProgress::Untouched,
        )?;
        consume_execution_resource(
            options,
            Resource::Work,
            source_content_length
                .checked_add(fragment_len)
                .ok_or_else(|| SourceContentPublicationError::Unsupported {
                    reason: "source-content insertion work accounting overflow".to_string(),
                })?,
            SourceContentPublicationProgress::Untouched,
        )?;

        let fragment_bytes = fragment.bytes();
        let mut scan_cancellation = None;
        let streamed_result =
            source.with_verified_member_reader_abortable(constants::ODF_CONTENT, |reader| {
                let mut spliced = SplicedReader::new(
                    reader,
                    source_content_length,
                    decoded_offset,
                    fragment_bytes,
                );
                let _report = scan_content_reader(
                    &mut spliced,
                    limits,
                    options,
                    &mut scan_cancellation,
                    &mut validator,
                )?;
                spliced.finish().map_err(|what| {
                    Error::InvalidFormat(format!(
                        "content.xml insertion stream proof failed: {what}"
                    ))
                })
            });
        let streamed = match streamed_result {
            Ok(streamed) => streamed,
            Err(error) => {
                // A source/archive failure is primary even when the callback
                // had already returned the cancellation sentinel.  Only a
                // callback-only sentinel is converted back to the typed
                // publication cancellation error.
                let source_primary = error.core().is_some();
                let callback_cancelled = !source_primary
                    && matches!(
                        error.callback(),
                        Some(Error::Other(message))
                            if message.as_str() == SOURCE_CONTENT_SCAN_CANCELLED_MESSAGE
                    );
                if callback_cancelled {
                    if let Some(error) = scan_cancellation.take() {
                        return Err(error);
                    }
                }
                return Err(map_member_prepare_error(
                    error,
                    SourceContentPublicationProgress::Untouched,
                ));
            },
        };
        if let Some(error) = scan_cancellation {
            return Err(error);
        }

        source.ensure_current_for_publication().map_err(|error| {
            super::map_core_error(error, SourceContentPublicationProgress::Untouched)
        })?;
        check_cancellation(options, SourceContentPublicationProgress::Untouched)?;
        if streamed.source_content_length != source_content_length
            || streamed.target_content_length != target_content_length
        {
            return Err(SourceContentPublicationError::Core(Error::InvalidFormat(
                "content.xml insertion stream length proof disagrees with source metadata"
                    .to_string(),
            )));
        }

        Ok(Self {
            source_version: source.source_version_snapshot(),
            source_length: source.len(),
            source,
            decoded_offset,
            fragment,
            _retained_memory: retained_memory,
            source_content_sha256: streamed.source_sha256,
            target_content_sha256: streamed.target_sha256,
            source_content_length,
            target_content_length,
        })
    }

    /// Return the prepared source version.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Return the source archive length captured by the plan.
    #[must_use]
    pub const fn source_length(&self) -> u64 {
        self.source_length
    }

    /// Return the decoded insertion offset.
    #[must_use]
    pub const fn decoded_offset(&self) -> u64 {
        self.decoded_offset
    }

    /// Return the source `content.xml` decoded length.
    #[must_use]
    pub const fn source_content_length(&self) -> u64 {
        self.source_content_length
    }

    /// Return the candidate `content.xml` decoded length.
    #[must_use]
    pub const fn target_content_length(&self) -> u64 {
        self.target_content_length
    }

    /// Return the source `content.xml` SHA-256 proof.
    #[must_use]
    pub const fn source_content_sha256(&self) -> [u8; 32] {
        self.source_content_sha256
    }

    /// Return the candidate `content.xml` SHA-256 proof.
    #[must_use]
    pub const fn target_content_sha256(&self) -> [u8; 32] {
        self.target_content_sha256
    }

    /// Publish the prepared insertion to a caller-owned sequential sink.
    ///
    /// Both ZIP producer passes open a fresh verified source-member reader and
    /// verify the source/candidate proof before the pass returns.  A sink or
    /// cancellation callback error aborts the source reader without draining
    /// it, and no automatic retry is attempted.
    ///
    /// The plan's retained-fragment reservation remains live through its
    /// lifetime. Publication acquires an independent reservation for that
    /// same retained capacity under the supplied options because the options
    /// may use an unrelated execution budget; a same-budget call therefore
    /// conservatively charges the capacity twice rather than assuming budget
    /// identity that this API cannot prove.
    pub fn write_to<W, O>(
        &self,
        writer: W,
        options: O,
    ) -> std::result::Result<SourceContentPublicationReport, SourceContentInsertionError>
    where
        W: Write,
        O: Borrow<SourceContentPublicationOptions>,
    {
        let options = options.borrow();
        let result = self.write_to_inner(writer, options);
        match result {
            Ok(report) => Ok(report),
            Err(error) => {
                if let Err(source) = reconcile_source_state(self.source.as_ref(), error.progress())
                {
                    return Err(SourceContentInsertionError::Publication {
                        progress: source.progress(),
                        source,
                    });
                }
                Err(error)
            },
        }
    }

    fn write_to_inner<W: Write>(
        &self,
        writer: W,
        options: &SourceContentPublicationOptions,
    ) -> std::result::Result<SourceContentPublicationReport, SourceContentInsertionError> {
        check_cancellation(options, SourceContentPublicationProgress::Untouched)
            .map_err(SourceContentInsertionError::from)?;
        let retained_fragment_bytes =
            retained_fragment_bytes(&self.fragment).map_err(SourceContentInsertionError::from)?;
        // Keep a publication-local charge even though the plan retains the
        // preparation lease.  The caller may supply an unrelated execution
        // context here; without a budget-identity relation, charging both is
        // the conservative choice and cannot under-account the retained Vec.
        let _publication_retained_memory = reserve_memory(
            options,
            retained_fragment_bytes,
            SourceContentPublicationProgress::Untouched,
        )
        .map_err(SourceContentInsertionError::from)?;
        self.source
            .ensure_current_for_publication()
            .map_err(|error| {
                SourceContentInsertionError::from(super::map_core_error(
                    error,
                    SourceContentPublicationProgress::Untouched,
                ))
            })?;

        let _input = self
            .source
            .begin_publication_input_accounting(options.execution_context())
            .map_err(|source| SourceContentInsertionError::Publication {
                progress: SourceContentPublicationProgress::Untouched,
                source: SourceContentPublicationError::Allocation {
                    resource: "source-content insertion input-accounting stack",
                    source,
                },
            })?;
        let fragment_len = u64::try_from(self.fragment.bytes().len()).map_err(|_| {
            SourceContentInsertionError::from(SourceContentPublicationError::Unsupported {
                reason: "authored insertion fragment length exceeds u64".to_string(),
            })
        })?;
        validate_replacement_limit(fragment_len, options)
            .map_err(SourceContentInsertionError::from)?;
        validate_replacement_limit(self.target_content_length, options)
            .map_err(SourceContentInsertionError::from)?;
        let replay_memory = insertion_transient_memory_requirement(
            self.source.publication_metadata_bytes(),
            u64::try_from(self.source.publication_entry_count()).unwrap_or(u64::MAX),
            0,
        )
        .map_err(SourceContentInsertionError::from)?;
        let _memory = reserve_memory(
            options,
            replay_memory,
            SourceContentPublicationProgress::Untouched,
        )
        .map_err(SourceContentInsertionError::from)?;
        let entry_count = u64::try_from(self.source.publication_entry_count()).unwrap_or(u64::MAX);
        let object_work = entry_count.checked_add(2).ok_or_else(|| {
            SourceContentInsertionError::from(SourceContentPublicationError::Unsupported {
                reason: "source-content insertion object accounting overflow".to_string(),
            })
        })?;
        consume_execution_resource(
            options,
            Resource::Objects,
            object_work,
            SourceContentPublicationProgress::Untouched,
        )
        .map_err(SourceContentInsertionError::from)?;
        let replay_work = self
            .target_content_length
            .checked_mul(2)
            .and_then(|value| value.checked_add(self.source.len()))
            .ok_or_else(|| {
                SourceContentInsertionError::from(SourceContentPublicationError::Unsupported {
                    reason: "source-content insertion replay work overflows u64".to_string(),
                })
            })?;
        consume_execution_resource(
            options,
            Resource::Work,
            replay_work,
            SourceContentPublicationProgress::Untouched,
        )
        .map_err(SourceContentInsertionError::from)?;

        let mut scratch = Vec::new();
        scratch
            .try_reserve_exact(soapberry_zip::RECOMMENDED_BUFFER_SIZE)
            .map_err(|source| SourceContentInsertionError::Publication {
                progress: SourceContentPublicationProgress::Untouched,
                source: SourceContentPublicationError::Allocation {
                    resource: "source-content insertion preservation index",
                    source,
                },
            })?;
        scratch.resize(soapberry_zip::RECOMMENDED_BUFFER_SIZE, 0);
        let index = self
            .source
            .preservation_index(&mut scratch)
            .map_err(|error| SourceContentInsertionError::Publication {
                progress: SourceContentPublicationProgress::Untouched,
                source: super::map_preflight_zip_error(error),
            })?;
        self.source
            .ensure_current_for_publication()
            .map_err(|error| {
                SourceContentInsertionError::from(super::map_core_error(
                    error,
                    SourceContentPublicationProgress::Untouched,
                ))
            })?;
        preflight_insertion_source(self.source.as_ref(), &index, options)
            .map_err(SourceContentInsertionError::from)?;

        if options.verify_payloads() {
            verify_untouched_payloads(self.source.as_ref(), options)
                .map_err(SourceContentInsertionError::from)?;
        }

        let content_entry = index
            .entries()
            .iter()
            .find(|entry| entry.raw_name_bytes() == constants::ODF_CONTENT.as_bytes())
            .ok_or_else(|| {
                SourceContentInsertionError::from(SourceContentPublicationError::Unsupported {
                    reason: "source archive has no canonical content.xml member".to_string(),
                })
            })?;
        let compression = match content_entry.compression_method() {
            CompressionMethod::Store | CompressionMethod::Deflate => {
                content_entry.compression_method()
            },
            other => {
                return Err(SourceContentInsertionError::from(
                    SourceContentPublicationError::Unsupported {
                        reason: format!("content.xml uses unsupported ZIP compression {other:?}"),
                    },
                ));
            },
        };

        let expected_version = self.source_version;
        let mut state = OutputState::new();
        let checked = CheckedSink {
            inner: writer,
            package: self.source.as_ref(),
            expected_version,
            options,
            state: &mut state,
        };
        let replay_limits = ReplayLimits::new(
            self.target_content_length.max(1),
            options.max_output_bytes(),
            options.max_output_bytes(),
        )
        .map_err(|error| SourceContentInsertionError::Publication {
            progress: SourceContentPublicationProgress::Untouched,
            source: super::map_preflight_zip_error(error),
        })?;
        let replay = index.write_replacing_with_replay(
            content_entry.id(),
            compression,
            replay_limits,
            checked,
            |sink| self.write_candidate_pass(sink, options),
        );
        let checked = match replay {
            Ok(checked) => checked,
            Err(source) => {
                let progress = state.progress();
                if let Some(publication) = map_replay_transport_error(&source, progress) {
                    return Err(SourceContentInsertionError::Transport {
                        progress,
                        publication,
                        replay: Box::new(source),
                    });
                }
                return Err(SourceContentInsertionError::Replay { progress, source });
            },
        };
        drop(checked);
        let progress = state.progress();
        check_cancellation(options, progress).map_err(SourceContentInsertionError::from)?;
        reconcile_source_state(self.source.as_ref(), progress)
            .map_err(SourceContentInsertionError::from)?;
        Ok(SourceContentPublicationReport::from_parts(
            progress.accepted(),
            false,
            expected_version,
        ))
    }

    fn write_candidate_pass(
        &self,
        sink: &mut dyn Write,
        options: &SourceContentPublicationOptions,
    ) -> std::result::Result<(), SourceContentInsertionCallbackError> {
        check_callback_cancellation(options)?;
        self.source
            .ensure_current_for_publication()
            .map_err(|error| {
                SourceContentInsertionCallbackError::Member(Box::new(
                    SourceMemberReaderError::Core {
                        error,
                        callback_error: None,
                    },
                ))
            })?;
        let fragment = self.fragment.bytes();
        let result =
            self.source
                .with_verified_member_reader_abortable(constants::ODF_CONTENT, |reader| {
                    let mut spliced = SplicedReader::new(
                        reader,
                        self.source_content_length,
                        self.decoded_offset,
                        fragment,
                    );
                    let mut buffer = [0_u8; COPY_CHUNK_SIZE];
                    loop {
                        check_callback_cancellation(options)?;
                        let read = spliced
                            .read(&mut buffer)
                            .map_err(SourceContentInsertionCallbackError::Io)?;
                        if read == 0 {
                            break;
                        }
                        sink.write_all(&buffer[..read])
                            .map_err(SourceContentInsertionCallbackError::Io)?;
                    }
                    let proof = spliced.finish().map_err(|what| {
                        SourceContentInsertionCallbackError::Verification { what }
                    })?;
                    if proof.source_sha256 != self.source_content_sha256 {
                        return Err(SourceContentInsertionCallbackError::Verification {
                            what: "source content SHA-256",
                        });
                    }
                    if proof.target_sha256 != self.target_content_sha256 {
                        return Err(SourceContentInsertionCallbackError::Verification {
                            what: "candidate content SHA-256",
                        });
                    }
                    if proof.source_content_length != self.source_content_length
                        || proof.target_content_length != self.target_content_length
                    {
                        return Err(SourceContentInsertionCallbackError::Verification {
                            what: "candidate content length",
                        });
                    }
                    Ok(())
                });
        match result {
            Ok(()) => {
                self.source
                    .ensure_current_for_publication()
                    .map_err(|error| {
                        SourceContentInsertionCallbackError::Member(Box::new(
                            SourceMemberReaderError::Core {
                                error,
                                callback_error: None,
                            },
                        ))
                    })?;
                Ok(())
            },
            Err(error) => Err(SourceContentInsertionCallbackError::Member(Box::new(error))),
        }
    }
}

#[derive(Debug)]
struct StreamProof {
    source_sha256: [u8; 32],
    target_sha256: [u8; 32],
    source_content_length: u64,
    target_content_length: u64,
}

/// A bounded sequential view of `prefix + fragment + suffix`.
struct SplicedReader<'source, 'fragment> {
    source: &'source mut dyn BufRead,
    expected_source_length: u64,
    prefix_remaining: u64,
    fragment: &'fragment [u8],
    fragment_position: usize,
    source_sha256: Sha256,
    target_sha256: Sha256,
    source_content_length: u64,
    target_content_length: u64,
    pending: [u8; COPY_CHUNK_SIZE],
    pending_start: usize,
    pending_end: usize,
    source_eof: bool,
}

impl<'source, 'fragment> SplicedReader<'source, 'fragment> {
    fn new(
        source: &'source mut dyn BufRead,
        expected_source_length: u64,
        offset: u64,
        fragment: &'fragment [u8],
    ) -> Self {
        Self {
            source,
            expected_source_length,
            prefix_remaining: offset,
            fragment,
            fragment_position: 0,
            source_sha256: Sha256::new(),
            target_sha256: Sha256::new(),
            source_content_length: 0,
            target_content_length: 0,
            pending: [0; COPY_CHUNK_SIZE],
            pending_start: 0,
            pending_end: 0,
            source_eof: false,
        }
    }

    fn append_source_pending(&mut self, length: usize) -> io::Result<()> {
        let bytes = &self.pending[..length];
        let source_length = self
            .source_content_length
            .checked_add(u64::try_from(bytes.len()).map_err(|_| {
                io::Error::new(io::ErrorKind::InvalidData, "source-content length overflow")
            })?)
            .ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidData, "source-content length overflow")
            })?;
        if source_length > self.expected_source_length {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "verified source content exceeded declared length",
            ));
        }
        self.source_sha256.update(bytes);
        self.target_sha256.update(bytes);
        self.source_content_length = source_length;
        self.target_content_length = self
            .target_content_length
            .checked_add(u64::try_from(bytes.len()).map_err(|_| {
                io::Error::new(io::ErrorKind::InvalidData, "candidate length overflow")
            })?)
            .ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidData, "candidate length overflow")
            })?;
        Ok(())
    }

    fn append_fragment_pending(&mut self, length: usize) -> io::Result<()> {
        let bytes = &self.pending[..length];
        self.target_sha256.update(bytes);
        self.target_content_length = self
            .target_content_length
            .checked_add(u64::try_from(bytes.len()).map_err(|_| {
                io::Error::new(io::ErrorKind::InvalidData, "candidate length overflow")
            })?)
            .ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidData, "candidate length overflow")
            })?;
        Ok(())
    }

    fn refill(&mut self) -> io::Result<()> {
        self.pending_start = 0;
        self.pending_end = 0;
        if self.prefix_remaining != 0 {
            let request = usize::try_from(self.prefix_remaining)
                .unwrap_or(COPY_CHUNK_SIZE)
                .min(COPY_CHUNK_SIZE);
            let read = self.source.read(&mut self.pending[..request])?;
            if read == 0 {
                return Err(io::Error::new(
                    io::ErrorKind::UnexpectedEof,
                    "source content ended before insertion offset",
                ));
            }
            self.prefix_remaining -= u64::try_from(read).map_err(|_| {
                io::Error::new(io::ErrorKind::InvalidData, "source-content prefix overflow")
            })?;
            self.append_source_pending(read)?;
            self.pending_end = read;
            return Ok(());
        }
        if self.fragment_position < self.fragment.len() {
            let remaining = self.fragment.len() - self.fragment_position;
            let count = remaining.min(COPY_CHUNK_SIZE);
            let end = self.fragment_position + count;
            self.pending[..count].copy_from_slice(&self.fragment[self.fragment_position..end]);
            self.fragment_position = end;
            self.append_fragment_pending(count)?;
            self.pending_end = count;
            return Ok(());
        }
        if self.source_content_length == self.expected_source_length {
            self.source_eof = true;
            return Ok(());
        }
        if self.source_eof {
            return Ok(());
        }
        let remaining = self.expected_source_length - self.source_content_length;
        let request = usize::try_from(remaining)
            .unwrap_or(COPY_CHUNK_SIZE)
            .min(COPY_CHUNK_SIZE);
        let read = self.source.read(&mut self.pending[..request])?;
        if read == 0 {
            self.source_eof = true;
            return Ok(());
        }
        self.append_source_pending(read)?;
        self.pending_end = read;
        Ok(())
    }

    fn finish(self) -> std::result::Result<StreamProof, &'static str> {
        if self.prefix_remaining != 0 || self.source_content_length != self.expected_source_length {
            return Err("source content length");
        }
        let expected_target = self
            .expected_source_length
            .checked_add(u64::try_from(self.fragment.len()).unwrap_or(u64::MAX))
            .ok_or("candidate content length")?;
        if self.target_content_length != expected_target {
            return Err("candidate content length");
        }
        Ok(StreamProof {
            source_sha256: self.source_sha256.finalize().into(),
            target_sha256: self.target_sha256.finalize().into(),
            source_content_length: self.source_content_length,
            target_content_length: self.target_content_length,
        })
    }
}

impl Read for SplicedReader<'_, '_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let mut written = 0usize;
        while written < output.len() {
            let available = self.fill_buf()?;
            if available.is_empty() {
                break;
            }
            let count = available.len().min(output.len() - written);
            output[written..written + count].copy_from_slice(&available[..count]);
            self.consume(count);
            written += count;
        }
        Ok(written)
    }
}

impl BufRead for SplicedReader<'_, '_> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.pending_start == self.pending_end && !self.source_eof {
            self.refill()?;
        }
        Ok(&self.pending[self.pending_start..self.pending_end])
    }

    fn consume(&mut self, amount: usize) {
        let available = self.pending_end.saturating_sub(self.pending_start);
        self.pending_start += amount.min(available);
    }
}

fn map_member_prepare_error(
    error: SourceMemberReaderError<Error>,
    progress: SourceContentPublicationProgress,
) -> SourceContentPublicationError {
    match error {
        SourceMemberReaderError::Core { error, .. } | SourceMemberReaderError::Callback(error) => {
            super::map_core_error(error, progress)
        },
    }
}

fn check_callback_cancellation(
    options: &SourceContentPublicationOptions,
) -> std::result::Result<(), SourceContentInsertionCallbackError> {
    if options
        .cancellation()
        .is_some_and(litchi_core::CancellationToken::is_cancelled)
    {
        return Err(SourceContentInsertionCallbackError::Cancelled);
    }
    if let Some(execution) = options.execution_context() {
        execution
            .check()
            .map_err(SourceContentInsertionCallbackError::Execution)?;
    }
    Ok(())
}

fn map_replay_transport_error(
    error: &ReplayPublicationError<SourceContentInsertionCallbackError>,
    progress: SourceContentPublicationProgress,
) -> Option<SourceContentPublicationError> {
    match error {
        ReplayPublicationError::Sink {
            source,
            callback_error,
            ..
        } => map_replay_io_marker(source, progress).or_else(|| {
            callback_error
                .as_ref()
                .and_then(|error| map_callback_transport_marker(error, progress))
        }),
        ReplayPublicationError::Archive {
            source,
            callback_error,
            ..
        } => map_replay_archive_marker(source, progress).or_else(|| {
            callback_error
                .as_ref()
                .and_then(|error| map_callback_transport_marker(error, progress))
        }),
        ReplayPublicationError::Limit {
            actual, maximum, ..
        } => Some(SourceContentPublicationError::LimitExceeded {
            progress,
            actual: *actual,
            maximum: *maximum,
        }),
        ReplayPublicationError::Callback { source, .. } => {
            map_callback_transport_marker(source, progress)
        },
        ReplayPublicationError::NonDeterministic { .. } => None,
        _ => None,
    }
}

fn map_replay_archive_marker(
    error: &soapberry_zip::Error,
    progress: SourceContentPublicationProgress,
) -> Option<SourceContentPublicationError> {
    match error.kind() {
        soapberry_zip::ErrorKind::IO(source) | soapberry_zip::ErrorKind::Io(source) => {
            map_replay_io_marker(source, progress)
        },
        soapberry_zip::ErrorKind::Cancelled => {
            Some(SourceContentPublicationError::Cancelled { progress })
        },
        soapberry_zip::ErrorKind::LimitExceeded {
            actual, maximum, ..
        } => Some(SourceContentPublicationError::LimitExceeded {
            progress,
            actual: *actual,
            maximum: *maximum,
        }),
        _ => None,
    }
}

fn map_callback_transport_marker(
    error: &SourceContentInsertionCallbackError,
    progress: SourceContentPublicationProgress,
) -> Option<SourceContentPublicationError> {
    match error {
        SourceContentInsertionCallbackError::Cancelled => {
            Some(SourceContentPublicationError::Cancelled { progress })
        },
        SourceContentInsertionCallbackError::Execution(source) => {
            Some(super::map_execution_error(source.clone(), progress))
        },
        SourceContentInsertionCallbackError::Io(source) => map_replay_io_marker(source, progress),
        SourceContentInsertionCallbackError::Member(source) => {
            map_member_transport_marker(source, progress)
        },
        SourceContentInsertionCallbackError::Verification { .. } => None,
    }
}

fn map_member_transport_marker(
    error: &SourceMemberReaderError<SourceContentInsertionCallbackError>,
    progress: SourceContentPublicationProgress,
) -> Option<SourceContentPublicationError> {
    match error {
        SourceMemberReaderError::Core { error, .. } => map_core_transport_marker(error, progress),
        SourceMemberReaderError::Callback(error) => map_callback_transport_marker(error, progress),
    }
}

fn map_core_transport_marker(
    error: &Error,
    progress: SourceContentPublicationProgress,
) -> Option<SourceContentPublicationError> {
    match error {
        Error::SourceChanged { expected, observed } => {
            Some(SourceContentPublicationError::SourceChanged {
                expected: *expected,
                observed: *observed,
                progress,
            })
        },
        Error::Io(source) => map_replay_io_marker(source, progress),
        _ => None,
    }
}

fn map_replay_io_marker(
    error: &io::Error,
    progress: SourceContentPublicationProgress,
) -> Option<SourceContentPublicationError> {
    if let Some(failure) = error
        .get_ref()
        .and_then(|source| source.downcast_ref::<super::PublicationIoFailure>())
    {
        return Some(super::map_publication_io_failure(failure, progress));
    }
    if let Some(changed) = error
        .get_ref()
        .and_then(|source| source.downcast_ref::<super::SourceChangedIo>())
    {
        return Some(SourceContentPublicationError::SourceChanged {
            expected: changed.expected,
            observed: changed.observed,
            progress,
        });
    }
    if let Some(source) = error
        .get_ref()
        .and_then(|source| source.downcast_ref::<super::SourceExecutionIo>())
    {
        return Some(super::map_execution_error(source.source.clone(), progress));
    }
    if let Some(source) = error
        .get_ref()
        .and_then(|source| source.downcast_ref::<super::SourceReadIo>())
    {
        return map_replay_io_marker(&source.source, progress).or_else(|| {
            Some(SourceContentPublicationError::Source {
                progress,
                source: Error::Io(copy_io_error(&source.source)),
            })
        });
    }
    if let Some(source) = error
        .get_ref()
        .and_then(|source| source.downcast_ref::<super::SinkWriteIo>())
    {
        return Some(SourceContentPublicationError::Sink {
            progress,
            source: copy_io_error(&source.source),
        });
    }
    None
}

fn copy_io_error(error: &io::Error) -> io::Error {
    io::Error::new(error.kind(), error.to_string())
}

fn retained_fragment_bytes(
    fragment: &AuthoredXmlFragment,
) -> std::result::Result<u64, SourceContentPublicationError> {
    u64::try_from(fragment.retention_bytes()).map_err(|_| {
        SourceContentPublicationError::Unsupported {
            reason: "authored insertion retained capacity exceeds u64".to_string(),
        }
    })
}

fn insertion_transient_memory_requirement(
    metadata: u64,
    entries: u64,
    xml_memory: u64,
) -> std::result::Result<u64, SourceContentPublicationError> {
    SOURCE_CONTENT_OPERATION_BUFFER_BYTES
        .checked_add(xml_memory)
        .and_then(|amount| amount.checked_add(metadata.checked_mul(2)?))
        .and_then(|amount| amount.checked_add(entries.checked_mul(INSERTION_OBJECT_BYTES)?))
        .ok_or_else(|| SourceContentPublicationError::Unsupported {
            reason: "source-content insertion memory accounting overflow".to_string(),
        })
}

fn validate_replacement_limit(
    observed: u64,
    options: &SourceContentPublicationOptions,
) -> std::result::Result<(), SourceContentPublicationError> {
    if observed > options.max_replacement_bytes() {
        return Err(SourceContentPublicationError::LimitExceeded {
            progress: SourceContentPublicationProgress::Untouched,
            actual: observed,
            maximum: options.max_replacement_bytes(),
        });
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
        let size = package.publication_member_size(path).map_err(|error| {
            super::map_core_error(error, SourceContentPublicationProgress::Untouched)
        })?;
        consume_execution_resource(
            options,
            Resource::Work,
            size,
            SourceContentPublicationProgress::Untouched,
        )?;
        let result = package.with_verified_member_reader_abortable(path, |reader| {
            let mut buffer = [0_u8; COPY_CHUNK_SIZE];
            loop {
                check_cancellation(options, SourceContentPublicationProgress::Untouched)
                    .map_err(VerifyCallbackError::Publication)?;
                let read = reader.read(&mut buffer).map_err(VerifyCallbackError::Io)?;
                if read == 0 {
                    break;
                }
            }
            Ok::<(), VerifyCallbackError>(())
        });
        match result {
            Ok(()) => {},
            Err(SourceMemberReaderError::Core { error, .. }) => {
                return Err(super::map_core_error(
                    error,
                    SourceContentPublicationProgress::Untouched,
                ));
            },
            Err(SourceMemberReaderError::Callback(VerifyCallbackError::Publication(error))) => {
                return Err(error);
            },
            Err(SourceMemberReaderError::Callback(VerifyCallbackError::Io(error))) => {
                return Err(SourceContentPublicationError::Source {
                    progress: SourceContentPublicationProgress::Untouched,
                    source: Error::Io(error),
                });
            },
        }
    }
    Ok(())
}

#[derive(Debug)]
enum VerifyCallbackError {
    Publication(SourceContentPublicationError),
    Io(io::Error),
}

impl fmt::Display for VerifyCallbackError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Publication(error) => error.fmt(formatter),
            Self::Io(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for VerifyCallbackError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Publication(error) => Some(error),
            Self::Io(error) => Some(error),
        }
    }
}
