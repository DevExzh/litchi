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
    source_remove_length: u64,
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
        validator: F,
    ) -> std::result::Result<Self, SourceContentPublicationError>
    where
        F: for<'event> FnMut(&'event XmlStreamEvent<'event>, &BindingTracker) -> Result<()>,
    {
        Self::prepare_replacement_with_validator(
            source,
            decoded_offset,
            0,
            input_fragment,
            limits,
            options,
            validator,
        )
    }

    /// Prepare a bounded replacement of a source range with an audited
    /// fragment.  A non-empty `source_remove_length` enables inverse replay
    /// without materializing the complete source or candidate XML member.
    /// The source range and replacement are streamed together, while the
    /// source and candidate SHA-256 proofs are retained for publication.
    #[doc(hidden)]
    pub fn prepare_replacement_with_validator<F>(
        source: Arc<SourceBackedPackage>,
        decoded_offset: u64,
        source_remove_length: u64,
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
        let source_end = decoded_offset
            .checked_add(source_remove_length)
            .ok_or_else(|| {
                SourceContentPublicationError::Core(Error::InvalidFormat(
                    "content.xml replacement range overflows decoded source length".to_string(),
                ))
            })?;
        if source_end > source_content_length {
            return Err(SourceContentPublicationError::Core(Error::InvalidFormat(
                "content.xml replacement range exceeds decoded source length".to_string(),
            )));
        }
        let target_content_length = source_content_length
            .checked_sub(source_remove_length)
            .and_then(|length| length.checked_add(fragment_len))
            .ok_or_else(|| SourceContentPublicationError::Unsupported {
                reason: "content.xml replacement target length overflows u64".to_string(),
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
                    source_remove_length,
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
            source_remove_length,
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

    /// Return the number of source bytes removed before the fragment is
    /// inserted.  Zero denotes the ordinary append/insertion shape.
    #[must_use]
    pub const fn source_remove_length(&self) -> u64 {
        self.source_remove_length
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

    /// Return the bounded authored fragment retained by this plan.
    #[doc(hidden)]
    pub fn authored_fragment_bytes(&self) -> &[u8] {
        self.fragment.bytes()
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
                        self.source_remove_length,
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
    source_remove_length: u64,
    fragment: &'fragment [u8],
    phase: SplicePhase,
    source_sha256: Sha256,
    target_sha256: Sha256,
    source_content_length: u64,
    target_content_length: u64,
    pending: [u8; COPY_CHUNK_SIZE],
    pending_start: usize,
    pending_end: usize,
}

/// Which part of `prefix + fragment + suffix` the next refill produces:
///
/// ```text
/// Prefix { remaining } --source bytes copied--> Prefix { remaining - read }
/// Prefix { remaining: 0 } --> Removal { remaining: removed length }
/// Removal { remaining } --source bytes dropped--> Removal { remaining - read }
/// Removal { remaining: 0 } --> Fragment { position: 0 }
/// Fragment { position } --fragment bytes copied--> Fragment { position + count }
/// Fragment { position: fragment length } --> Suffix
/// Suffix --source bytes copied--> Suffix
/// Suffix --declared source length reached, or the source ended--> Done
/// ```
///
/// The phases only move forward, and a phase with nothing left is passed
/// without touching the source. `Prefix` and `Removal` fail when the source
/// ends early; the phase stays where it was, so a later refill asks the
/// source again. `Done` is final: the reader then reports end of input.
#[derive(Clone, Copy, PartialEq, Eq)]
enum SplicePhase {
    /// Source bytes before the insertion offset are still to be copied.
    Prefix {
        remaining: u64,
    },
    /// Replaced source bytes are still to be read and dropped.
    Removal {
        remaining: u64,
    },
    /// The authored fragment is copied from `position` on.
    Fragment {
        position: usize,
    },
    /// Source bytes after the replaced range are copied.
    Suffix,
    Done,
}

impl<'source, 'fragment> SplicedReader<'source, 'fragment> {
    fn new(
        source: &'source mut dyn BufRead,
        expected_source_length: u64,
        offset: u64,
        source_remove_length: u64,
        fragment: &'fragment [u8],
    ) -> Self {
        Self {
            source,
            expected_source_length,
            source_remove_length,
            fragment,
            phase: SplicePhase::Prefix { remaining: offset },
            source_sha256: Sha256::new(),
            target_sha256: Sha256::new(),
            source_content_length: 0,
            target_content_length: 0,
            pending: [0; COPY_CHUNK_SIZE],
            pending_start: 0,
            pending_end: 0,
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

    fn append_removed_source(&mut self, length: usize) -> io::Result<()> {
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
        self.source_content_length = source_length;
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
        loop {
            match self.phase {
                SplicePhase::Prefix { remaining: 0 } => {
                    self.phase = SplicePhase::Removal {
                        remaining: self.source_remove_length,
                    };
                },
                SplicePhase::Prefix { remaining } => {
                    let request = usize::try_from(remaining)
                        .unwrap_or(COPY_CHUNK_SIZE)
                        .min(COPY_CHUNK_SIZE);
                    let read = self.source.read(&mut self.pending[..request])?;
                    if read == 0 {
                        return Err(io::Error::new(
                            io::ErrorKind::UnexpectedEof,
                            "source content ended before insertion offset",
                        ));
                    }
                    let consumed = u64::try_from(read).map_err(|_| {
                        io::Error::new(io::ErrorKind::InvalidData, "source-content prefix overflow")
                    })?;
                    self.phase = SplicePhase::Prefix {
                        remaining: remaining - consumed,
                    };
                    self.append_source_pending(read)?;
                    self.pending_end = read;
                    return Ok(());
                },
                SplicePhase::Removal { remaining: 0 } => {
                    self.phase = SplicePhase::Fragment { position: 0 };
                },
                SplicePhase::Removal { remaining } => {
                    let request = usize::try_from(remaining)
                        .unwrap_or(COPY_CHUNK_SIZE)
                        .min(COPY_CHUNK_SIZE);
                    let read = self.source.read(&mut self.pending[..request])?;
                    if read == 0 {
                        return Err(io::Error::new(
                            io::ErrorKind::UnexpectedEof,
                            "source content ended inside replacement range",
                        ));
                    }
                    let consumed = u64::try_from(read).map_err(|_| {
                        io::Error::new(
                            io::ErrorKind::InvalidData,
                            "source-content removal overflow",
                        )
                    })?;
                    self.phase = SplicePhase::Removal {
                        remaining: remaining - consumed,
                    };
                    self.append_removed_source(read)?;
                },
                SplicePhase::Fragment { position } if position < self.fragment.len() => {
                    let remaining = self.fragment.len() - position;
                    let count = remaining.min(COPY_CHUNK_SIZE);
                    let end = position + count;
                    self.pending[..count].copy_from_slice(&self.fragment[position..end]);
                    self.phase = SplicePhase::Fragment { position: end };
                    self.append_fragment_pending(count)?;
                    self.pending_end = count;
                    return Ok(());
                },
                SplicePhase::Fragment { .. } => self.phase = SplicePhase::Suffix,
                SplicePhase::Suffix => {
                    if self.source_content_length == self.expected_source_length {
                        self.phase = SplicePhase::Done;
                        return Ok(());
                    }
                    let remaining = self.expected_source_length - self.source_content_length;
                    let request = usize::try_from(remaining)
                        .unwrap_or(COPY_CHUNK_SIZE)
                        .min(COPY_CHUNK_SIZE);
                    let read = self.source.read(&mut self.pending[..request])?;
                    if read == 0 {
                        self.phase = SplicePhase::Done;
                        return Ok(());
                    }
                    self.append_source_pending(read)?;
                    self.pending_end = read;
                    return Ok(());
                },
                SplicePhase::Done => return Ok(()),
            }
        }
    }

    fn finish(self) -> std::result::Result<StreamProof, &'static str> {
        // Whether source bytes in front of the fragment were never read:
        // the rest of the prefix or of the replaced range.
        let source_before_fragment_unread = match self.phase {
            SplicePhase::Prefix { remaining } => remaining != 0 || self.source_remove_length != 0,
            SplicePhase::Removal { remaining } => remaining != 0,
            SplicePhase::Fragment { .. } | SplicePhase::Suffix | SplicePhase::Done => false,
        };
        if source_before_fragment_unread
            || self.source_content_length != self.expected_source_length
        {
            return Err("source content length");
        }
        let expected_target = self
            .expected_source_length
            .checked_sub(self.source_remove_length)
            .and_then(|length| {
                length.checked_add(u64::try_from(self.fragment.len()).unwrap_or(u64::MAX))
            })
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
        if self.pending_start == self.pending_end && self.phase != SplicePhase::Done {
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
        soapberry_zip::ErrorKind::IO(source) => map_replay_io_marker(source, progress),
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

#[cfg(test)]
mod tests {
    use super::{COPY_CHUNK_SIZE, SplicedReader};
    use std::cell::RefCell;
    use std::io::{self, BufRead, Read};

    /// What the scripted source does on one `read` call. Once the script is
    /// used up the source fills every request as far as its data allows.
    #[derive(Clone, Copy)]
    enum Step {
        /// Return at most this many bytes; zero reports end of input.
        Give(usize),
        Fail(io::ErrorKind),
    }

    struct ScriptedSource<'log> {
        data: Vec<u8>,
        position: usize,
        script: Vec<Step>,
        next: usize,
        /// One entry per `read` call: `request>returned`, or `request!` for
        /// a scripted failure.
        log: &'log RefCell<Vec<String>>,
    }

    impl Read for ScriptedSource<'_> {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            let step = self.script.get(self.next).copied();
            self.next += 1;
            let limit = match step {
                Some(Step::Fail(kind)) => {
                    self.log.borrow_mut().push(format!("{}!", output.len()));
                    return Err(io::Error::new(kind, "scripted source failure"));
                },
                Some(Step::Give(limit)) => limit,
                None => usize::MAX,
            };
            let count = limit.min(output.len()).min(self.data.len() - self.position);
            output[..count].copy_from_slice(&self.data[self.position..self.position + count]);
            self.position += count;
            self.log
                .borrow_mut()
                .push(format!("{}>{count}", output.len()));
            Ok(count)
        }
    }

    impl BufRead for ScriptedSource<'_> {
        fn fill_buf(&mut self) -> io::Result<&[u8]> {
            Ok(&self.data[self.position..])
        }

        fn consume(&mut self, amount: usize) {
            self.position = (self.position + amount).min(self.data.len());
        }
    }

    struct Case {
        data: Vec<u8>,
        script: Vec<Step>,
        expected_source_length: u64,
        offset: u64,
        remove: u64,
        fragment: Vec<u8>,
    }

    impl Case {
        fn digits(offset: u64, remove: u64, fragment: &str) -> Self {
            Self {
                data: b"0123456789".to_vec(),
                script: Vec::new(),
                expected_source_length: 10,
                offset,
                remove,
                fragment: fragment.as_bytes().to_vec(),
            }
        }

        fn script(mut self, script: &[Step]) -> Self {
            self.script = script.to_vec();
            self
        }

        fn data(mut self, data: &[u8]) -> Self {
            self.data = data.to_vec();
            self
        }

        fn expected(mut self, length: u64) -> Self {
            self.expected_source_length = length;
            self
        }
    }

    #[derive(Clone, Copy)]
    enum Op {
        Read(usize),
        FillBuf,
        Consume(usize),
    }

    fn render(bytes: &[u8]) -> String {
        if bytes.len() <= 32 {
            format!("{:?}", String::from_utf8_lossy(bytes))
        } else {
            format!(
                "{} bytes ending {:?}",
                bytes.len(),
                String::from_utf8_lossy(&bytes[bytes.len() - 4..])
            )
        }
    }

    fn hex4(hash: [u8; 32]) -> String {
        hash[..4].iter().map(|byte| format!("{byte:02x}")).collect()
    }

    /// Runs `ops` and then `finish`; returns one line per operation and a
    /// last line for `finish`. Each line ends with the source reads the
    /// operation caused.
    fn transcript(case: &Case, ops: &[Op]) -> Vec<String> {
        let log = RefCell::new(Vec::new());
        let mut source = ScriptedSource {
            data: case.data.clone(),
            position: 0,
            script: case.script.clone(),
            next: 0,
            log: &log,
        };
        let mut spliced = SplicedReader::new(
            &mut source,
            case.expected_source_length,
            case.offset,
            case.remove,
            &case.fragment,
        );
        let mut lines = Vec::new();
        for op in ops {
            let outcome = match *op {
                Op::Read(size) => {
                    let mut buffer = vec![0u8; size];
                    match spliced.read(&mut buffer) {
                        Ok(count) => format!("read({size}) => {}", render(&buffer[..count])),
                        Err(error) => format!("read({size}) => {:?}: {error}", error.kind()),
                    }
                },
                Op::FillBuf => match spliced.fill_buf() {
                    Ok(bytes) => format!("fill_buf => {}", render(bytes)),
                    Err(error) => format!("fill_buf => {:?}: {error}", error.kind()),
                },
                Op::Consume(amount) => {
                    spliced.consume(amount);
                    format!("consume({amount})")
                },
            };
            let reads = log.borrow_mut().drain(..).collect::<Vec<_>>().join(" ");
            lines.push(format!("{outcome} [{reads}]"));
        }
        lines.push(match spliced.finish() {
            Ok(proof) => format!(
                "finish => source {} bytes {}, target {} bytes {}",
                proof.source_content_length,
                hex4(proof.source_sha256),
                proof.target_content_length,
                hex4(proof.target_sha256)
            ),
            Err(what) => format!("finish => Err({what:?})"),
        });
        lines
    }

    /// The `finish` outcome after each prefix of `ops`.
    fn finish_after_each(case: &Case, ops: &[Op]) -> Vec<String> {
        (0..=ops.len())
            .map(|count| {
                let lines = transcript(case, &ops[..count]);
                format!("after {count}: {}", lines[lines.len() - 1])
            })
            .collect()
    }

    fn assert_lines(actual: &[String], expected: &[&str]) {
        let listing: Vec<String> = actual.iter().map(|line| format!("{line:?},")).collect();
        assert!(
            actual
                .iter()
                .map(String::as_str)
                .eq(expected.iter().copied()),
            "observed lines:\n{}",
            listing.join("\n")
        );
    }

    #[test]
    fn splice_walks_prefix_removal_fragment_suffix() {
        use Op::Read;
        let case = Case::digits(3, 2, "AB");
        assert_lines(
            &transcript(&case, &[Read(0), Read(64), Read(64), Read(0), Read(64)]),
            &[
                "read(0) => \"\" []",
                "read(64) => \"012AB56789\" [3>3 2>2 5>5]",
                "read(64) => \"\" []",
                "read(0) => \"\" []",
                "read(64) => \"\" []",
                "finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
            ],
        );
        assert_lines(
            &transcript(&case, &[Read(4), Read(4), Read(4), Read(4), Read(4)]),
            &[
                "read(4) => \"012A\" [3>3 2>2]",
                "read(4) => \"B567\" [5>5]",
                "read(4) => \"89\" []",
                "read(4) => \"\" []",
                "read(4) => \"\" []",
                "finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
            ],
        );
    }

    #[test]
    fn splice_finish_outcome_after_every_step() {
        let case = Case::digits(3, 2, "AB").script(&[Step::Give(1); 16]);
        assert_lines(
            &finish_after_each(&case, &[Op::Read(1); 13]),
            &[
                "after 0: finish => Err(\"source content length\")",
                "after 1: finish => Err(\"source content length\")",
                "after 2: finish => Err(\"source content length\")",
                "after 3: finish => Err(\"source content length\")",
                "after 4: finish => Err(\"source content length\")",
                "after 5: finish => Err(\"source content length\")",
                "after 6: finish => Err(\"source content length\")",
                "after 7: finish => Err(\"source content length\")",
                "after 8: finish => Err(\"source content length\")",
                "after 9: finish => Err(\"source content length\")",
                "after 10: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
                "after 11: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
                "after 12: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
                "after 13: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
            ],
        );
        let ops = [Op::FillBuf, Op::Consume(64)].repeat(12);
        assert_lines(
            &finish_after_each(&case, &ops),
            &[
                "after 0: finish => Err(\"source content length\")",
                "after 1: finish => Err(\"source content length\")",
                "after 2: finish => Err(\"source content length\")",
                "after 3: finish => Err(\"source content length\")",
                "after 4: finish => Err(\"source content length\")",
                "after 5: finish => Err(\"source content length\")",
                "after 6: finish => Err(\"source content length\")",
                "after 7: finish => Err(\"source content length\")",
                "after 8: finish => Err(\"source content length\")",
                "after 9: finish => Err(\"source content length\")",
                "after 10: finish => Err(\"source content length\")",
                "after 11: finish => Err(\"source content length\")",
                "after 12: finish => Err(\"source content length\")",
                "after 13: finish => Err(\"source content length\")",
                "after 14: finish => Err(\"source content length\")",
                "after 15: finish => Err(\"source content length\")",
                "after 16: finish => Err(\"source content length\")",
                "after 17: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
                "after 18: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
                "after 19: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
                "after 20: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
                "after 21: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
                "after 22: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
                "after 23: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
                "after 24: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
            ],
        );
    }

    #[test]
    fn splice_short_source_reads_keep_phase_boundaries() {
        let case = Case::digits(3, 2, "AB").script(&[Step::Give(1); 16]);
        assert_lines(
            &transcript(&case, &[Op::Read(64), Op::Read(64)]),
            &[
                "read(64) => \"012AB56789\" [3>1 2>1 1>1 2>1 1>1 5>1 4>1 3>1 2>1 1>1]",
                "read(64) => \"\" []",
                "finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
            ],
        );
        let ops = [Op::FillBuf, Op::Consume(64)].repeat(12);
        assert_lines(
            &transcript(&case, &ops),
            &[
                "fill_buf => \"0\" [3>1]",
                "consume(64) []",
                "fill_buf => \"1\" [2>1]",
                "consume(64) []",
                "fill_buf => \"2\" [1>1]",
                "consume(64) []",
                "fill_buf => \"AB\" [2>1 1>1]",
                "consume(64) []",
                "fill_buf => \"5\" [5>1]",
                "consume(64) []",
                "fill_buf => \"6\" [4>1]",
                "consume(64) []",
                "fill_buf => \"7\" [3>1]",
                "consume(64) []",
                "fill_buf => \"8\" [2>1]",
                "consume(64) []",
                "fill_buf => \"9\" [1>1]",
                "consume(64) []",
                "fill_buf => \"\" []",
                "consume(64) []",
                "fill_buf => \"\" []",
                "consume(64) []",
                "fill_buf => \"\" []",
                "consume(64) []",
                "finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
            ],
        );
        let case = Case::digits(3, 2, "AB").script(&[Step::Give(2); 16]);
        let ops = [Op::FillBuf, Op::Consume(64)].repeat(8);
        assert_lines(
            &transcript(&case, &ops),
            &[
                "fill_buf => \"01\" [3>2]",
                "consume(64) []",
                "fill_buf => \"2\" [1>1]",
                "consume(64) []",
                "fill_buf => \"AB\" [2>2]",
                "consume(64) []",
                "fill_buf => \"56\" [5>2]",
                "consume(64) []",
                "fill_buf => \"78\" [3>2]",
                "consume(64) []",
                "fill_buf => \"9\" [1>1]",
                "consume(64) []",
                "fill_buf => \"\" []",
                "consume(64) []",
                "fill_buf => \"\" []",
                "consume(64) []",
                "finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
            ],
        );
    }

    #[test]
    fn splice_buffered_view_is_stable_until_consumed() {
        use Op::{Consume, FillBuf, Read};
        let case = Case::digits(3, 2, "AB");
        assert_lines(
            &transcript(
                &case,
                &[
                    Consume(5),
                    FillBuf,
                    FillBuf,
                    Consume(1),
                    FillBuf,
                    Read(1),
                    FillBuf,
                    Consume(64),
                    FillBuf,
                    Consume(0),
                    FillBuf,
                    Consume(1),
                    Read(3),
                    FillBuf,
                    Consume(64),
                    FillBuf,
                    FillBuf,
                    Consume(1),
                    FillBuf,
                ],
            ),
            &[
                "consume(5) []",
                "fill_buf => \"012\" [3>3]",
                "fill_buf => \"012\" []",
                "consume(1) []",
                "fill_buf => \"12\" []",
                "read(1) => \"1\" []",
                "fill_buf => \"2\" []",
                "consume(64) []",
                "fill_buf => \"AB\" [2>2]",
                "consume(0) []",
                "fill_buf => \"AB\" []",
                "consume(1) []",
                "read(3) => \"B56\" [5>5]",
                "fill_buf => \"789\" []",
                "consume(64) []",
                "fill_buf => \"\" []",
                "fill_buf => \"\" []",
                "consume(1) []",
                "fill_buf => \"\" []",
                "finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
            ],
        );
    }

    #[test]
    fn splice_source_errors_at_and_inside_each_phase_are_retried_in_place() {
        use Step::{Fail, Give};
        let kind = io::ErrorKind::Other;
        // Offset 4, removal 3: one failure before and one inside the prefix,
        // the removal and the suffix.
        let script = [
            Fail(kind),
            Give(2),
            Fail(kind),
            Give(2),
            Fail(kind),
            Give(1),
            Fail(kind),
            Give(2),
            Fail(kind),
            Give(1),
            Fail(io::ErrorKind::Interrupted),
            Give(2),
        ];
        let case = Case::digits(4, 3, "AB").script(&script);
        let ops = [Op::FillBuf, Op::Consume(64)].repeat(12);
        assert_lines(
            &transcript(&case, &ops),
            &[
                "fill_buf => Other: scripted source failure [4!]",
                "consume(64) []",
                "fill_buf => \"01\" [4>2]",
                "consume(64) []",
                "fill_buf => Other: scripted source failure [2!]",
                "consume(64) []",
                "fill_buf => \"23\" [2>2]",
                "consume(64) []",
                "fill_buf => Other: scripted source failure [3!]",
                "consume(64) []",
                "fill_buf => Other: scripted source failure [3>1 2!]",
                "consume(64) []",
                "fill_buf => \"AB\" [2>2]",
                "consume(64) []",
                "fill_buf => Other: scripted source failure [3!]",
                "consume(64) []",
                "fill_buf => \"7\" [3>1]",
                "consume(64) []",
                "fill_buf => Interrupted: scripted source failure [2!]",
                "consume(64) []",
                "fill_buf => \"89\" [2>2]",
                "consume(64) []",
                "fill_buf => \"\" []",
                "consume(64) []",
                "finish => source 10 bytes 84d89877, target 9 bytes 4c353d10",
            ],
        );
        assert_lines(
            &finish_after_each(&case, &ops),
            &[
                "after 0: finish => Err(\"source content length\")",
                "after 1: finish => Err(\"source content length\")",
                "after 2: finish => Err(\"source content length\")",
                "after 3: finish => Err(\"source content length\")",
                "after 4: finish => Err(\"source content length\")",
                "after 5: finish => Err(\"source content length\")",
                "after 6: finish => Err(\"source content length\")",
                "after 7: finish => Err(\"source content length\")",
                "after 8: finish => Err(\"source content length\")",
                "after 9: finish => Err(\"source content length\")",
                "after 10: finish => Err(\"source content length\")",
                "after 11: finish => Err(\"source content length\")",
                "after 12: finish => Err(\"source content length\")",
                "after 13: finish => Err(\"source content length\")",
                "after 14: finish => Err(\"source content length\")",
                "after 15: finish => Err(\"source content length\")",
                "after 16: finish => Err(\"source content length\")",
                "after 17: finish => Err(\"source content length\")",
                "after 18: finish => Err(\"source content length\")",
                "after 19: finish => Err(\"source content length\")",
                "after 20: finish => Err(\"source content length\")",
                "after 21: finish => source 10 bytes 84d89877, target 9 bytes 4c353d10",
                "after 22: finish => source 10 bytes 84d89877, target 9 bytes 4c353d10",
                "after 23: finish => source 10 bytes 84d89877, target 9 bytes 4c353d10",
                "after 24: finish => source 10 bytes 84d89877, target 9 bytes 4c353d10",
            ],
        );
        // `read` gathers several refills; a failure discards what the same
        // call had already gathered.
        assert_lines(
            &transcript(&case, &[Op::Read(64); 8]),
            &[
                "read(64) => Other: scripted source failure [4!]",
                "read(64) => Other: scripted source failure [4>2 2!]",
                "read(64) => Other: scripted source failure [2>2 3!]",
                "read(64) => Other: scripted source failure [3>1 2!]",
                "read(64) => Other: scripted source failure [2>2 3!]",
                "read(64) => Interrupted: scripted source failure [3>1 2!]",
                "read(64) => \"89\" [2>2]",
                "read(64) => \"\" []",
                "finish => source 10 bytes 84d89877, target 9 bytes 4c353d10",
            ],
        );
        assert_lines(
            &transcript(&case, &[Op::Read(1); 16]),
            &[
                "read(1) => Other: scripted source failure [4!]",
                "read(1) => \"0\" [4>2]",
                "read(1) => \"1\" []",
                "read(1) => Other: scripted source failure [2!]",
                "read(1) => \"2\" [2>2]",
                "read(1) => \"3\" []",
                "read(1) => Other: scripted source failure [3!]",
                "read(1) => Other: scripted source failure [3>1 2!]",
                "read(1) => \"A\" [2>2]",
                "read(1) => \"B\" []",
                "read(1) => Other: scripted source failure [3!]",
                "read(1) => \"7\" [3>1]",
                "read(1) => Interrupted: scripted source failure [2!]",
                "read(1) => \"8\" [2>2]",
                "read(1) => \"9\" []",
                "read(1) => \"\" []",
                "finish => source 10 bytes 84d89877, target 9 bytes 4c353d10",
            ],
        );
    }

    #[test]
    fn splice_zero_length_source_reads_in_each_phase() {
        use Step::Give;
        let ops = [Op::FillBuf, Op::Consume(64)].repeat(6);
        let prefix = Case::digits(3, 2, "AB").script(&[Give(1), Give(0), Give(0)]);
        assert_lines(
            &transcript(&prefix, &ops),
            &[
                "fill_buf => \"0\" [3>1]",
                "consume(64) []",
                "fill_buf => UnexpectedEof: source content ended before insertion offset [2>0]",
                "consume(64) []",
                "fill_buf => UnexpectedEof: source content ended before insertion offset [2>0]",
                "consume(64) []",
                "fill_buf => \"12\" [2>2]",
                "consume(64) []",
                "fill_buf => \"AB\" [2>2]",
                "consume(64) []",
                "fill_buf => \"56789\" [5>5]",
                "consume(64) []",
                "finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
            ],
        );
        let removal = Case::digits(3, 2, "AB").script(&[Give(3), Give(1), Give(0), Give(0)]);
        assert_lines(
            &transcript(&removal, &ops),
            &[
                "fill_buf => \"012\" [3>3]",
                "consume(64) []",
                "fill_buf => UnexpectedEof: source content ended inside replacement range [2>1 1>0]",
                "consume(64) []",
                "fill_buf => UnexpectedEof: source content ended inside replacement range [1>0]",
                "consume(64) []",
                "fill_buf => \"AB\" [1>1]",
                "consume(64) []",
                "fill_buf => \"56789\" [5>5]",
                "consume(64) []",
                "fill_buf => \"\" []",
                "consume(64) []",
                "finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
            ],
        );
        let suffix = Case::digits(3, 2, "AB").script(&[Give(3), Give(2), Give(1), Give(0)]);
        assert_lines(
            &transcript(&suffix, &ops),
            &[
                "fill_buf => \"012\" [3>3]",
                "consume(64) []",
                "fill_buf => \"AB\" [2>2]",
                "consume(64) []",
                "fill_buf => \"5\" [5>1]",
                "consume(64) []",
                "fill_buf => \"\" [4>0]",
                "consume(64) []",
                "fill_buf => \"\" []",
                "consume(64) []",
                "fill_buf => \"\" []",
                "consume(64) []",
                "finish => Err(\"source content length\")",
            ],
        );
        assert_lines(
            &transcript(&suffix, &[Op::Read(64); 3]),
            &[
                "read(64) => \"012AB5\" [3>3 2>2 5>1 4>0]",
                "read(64) => \"\" []",
                "read(64) => \"\" []",
                "finish => Err(\"source content length\")",
            ],
        );
    }

    #[test]
    fn splice_source_shorter_or_longer_than_declared() {
        let ops = [Op::Read(64); 3];
        let rows: Vec<String> = [
            ("ends in prefix", Case::digits(3, 2, "AB").data(b"01")),
            ("ends at offset", Case::digits(3, 2, "AB").data(b"012")),
            ("ends in removal", Case::digits(3, 2, "AB").data(b"0123")),
            (
                "ends after removal",
                Case::digits(3, 2, "AB").data(b"01234"),
            ),
            ("ends in suffix", Case::digits(3, 2, "AB").data(b"0123456")),
            ("empty source", Case::digits(3, 2, "AB").data(b"")),
            (
                "longer than declared",
                Case::digits(3, 2, "AB").data(b"0123456789xy"),
            ),
            (
                "offset beyond declared",
                Case::digits(3, 2, "AB").expected(2),
            ),
            (
                "removal beyond declared",
                Case::digits(3, 2, "AB").expected(4),
            ),
            (
                "declared ends at removal",
                Case::digits(3, 2, "AB").expected(5),
            ),
        ]
        .iter()
        .flat_map(|(label, case)| {
            transcript(case, &ops)
                .into_iter()
                .map(move |line| format!("{label}: {line}"))
        })
        .collect();
        assert_lines(
            &rows,
            &[
                "ends in prefix: read(64) => UnexpectedEof: source content ended before insertion offset [3>2 1>0]",
                "ends in prefix: read(64) => UnexpectedEof: source content ended before insertion offset [1>0]",
                "ends in prefix: read(64) => UnexpectedEof: source content ended before insertion offset [1>0]",
                "ends in prefix: finish => Err(\"source content length\")",
                "ends at offset: read(64) => UnexpectedEof: source content ended inside replacement range [3>3 2>0]",
                "ends at offset: read(64) => UnexpectedEof: source content ended inside replacement range [2>0]",
                "ends at offset: read(64) => UnexpectedEof: source content ended inside replacement range [2>0]",
                "ends at offset: finish => Err(\"source content length\")",
                "ends in removal: read(64) => UnexpectedEof: source content ended inside replacement range [3>3 2>1 1>0]",
                "ends in removal: read(64) => UnexpectedEof: source content ended inside replacement range [1>0]",
                "ends in removal: read(64) => UnexpectedEof: source content ended inside replacement range [1>0]",
                "ends in removal: finish => Err(\"source content length\")",
                "ends after removal: read(64) => \"012AB\" [3>3 2>2 5>0]",
                "ends after removal: read(64) => \"\" []",
                "ends after removal: read(64) => \"\" []",
                "ends after removal: finish => Err(\"source content length\")",
                "ends in suffix: read(64) => \"012AB56\" [3>3 2>2 5>2 3>0]",
                "ends in suffix: read(64) => \"\" []",
                "ends in suffix: read(64) => \"\" []",
                "ends in suffix: finish => Err(\"source content length\")",
                "empty source: read(64) => UnexpectedEof: source content ended before insertion offset [3>0]",
                "empty source: read(64) => UnexpectedEof: source content ended before insertion offset [3>0]",
                "empty source: read(64) => UnexpectedEof: source content ended before insertion offset [3>0]",
                "empty source: finish => Err(\"source content length\")",
                "longer than declared: read(64) => \"012AB56789\" [3>3 2>2 5>5]",
                "longer than declared: read(64) => \"\" []",
                "longer than declared: read(64) => \"\" []",
                "longer than declared: finish => source 10 bytes 84d89877, target 10 bytes 30c46a20",
                "offset beyond declared: read(64) => InvalidData: verified source content exceeded declared length [3>3]",
                "offset beyond declared: read(64) => \"AB\" [2>2]",
                "offset beyond declared: read(64) => \"\" []",
                "offset beyond declared: finish => source 2 bytes 86e50149, target 2 bytes 38164fbd",
                "removal beyond declared: read(64) => InvalidData: verified source content exceeded declared length [3>3 2>2]",
                "removal beyond declared: read(64) => \"AB5\" [1>1]",
                "removal beyond declared: read(64) => \"\" []",
                "removal beyond declared: finish => Err(\"candidate content length\")",
                "declared ends at removal: read(64) => \"012AB\" [3>3 2>2]",
                "declared ends at removal: read(64) => \"\" []",
                "declared ends at removal: read(64) => \"\" []",
                "declared ends at removal: finish => source 5 bytes c565fe03, target 5 bytes 61bc9aea",
            ],
        );
    }

    #[test]
    fn splice_empty_phases_are_skipped() {
        let ops = [Op::FillBuf, Op::Consume(64)].repeat(5);
        let rows: Vec<String> = [
            ("no prefix", Case::digits(0, 2, "AB")),
            ("no removal", Case::digits(3, 0, "AB")),
            ("no fragment", Case::digits(3, 2, "")),
            ("no suffix", Case::digits(8, 2, "AB")),
            ("insert at end", Case::digits(10, 0, "AB")),
            ("insert at start", Case::digits(0, 0, "AB")),
            ("replace everything", Case::digits(0, 10, "AB")),
            ("remove everything", Case::digits(0, 10, "")),
            ("nothing to do", Case::digits(0, 0, "")),
            (
                "empty source and fragment",
                Case::digits(0, 0, "").data(b"").expected(0),
            ),
            (
                "empty source",
                Case::digits(0, 0, "AB").data(b"").expected(0),
            ),
        ]
        .iter()
        .flat_map(|(label, case)| {
            transcript(case, &ops)
                .into_iter()
                .map(move |line| format!("{label}: {line}"))
        })
        .collect();
        assert_lines(
            &rows,
            &[
                "no prefix: fill_buf => \"AB\" [2>2]",
                "no prefix: consume(64) []",
                "no prefix: fill_buf => \"23456789\" [8>8]",
                "no prefix: consume(64) []",
                "no prefix: fill_buf => \"\" []",
                "no prefix: consume(64) []",
                "no prefix: fill_buf => \"\" []",
                "no prefix: consume(64) []",
                "no prefix: fill_buf => \"\" []",
                "no prefix: consume(64) []",
                "no prefix: finish => source 10 bytes 84d89877, target 10 bytes b38acc48",
                "no removal: fill_buf => \"012\" [3>3]",
                "no removal: consume(64) []",
                "no removal: fill_buf => \"AB\" []",
                "no removal: consume(64) []",
                "no removal: fill_buf => \"3456789\" [7>7]",
                "no removal: consume(64) []",
                "no removal: fill_buf => \"\" []",
                "no removal: consume(64) []",
                "no removal: fill_buf => \"\" []",
                "no removal: consume(64) []",
                "no removal: finish => source 10 bytes 84d89877, target 12 bytes 50716a22",
                "no fragment: fill_buf => \"012\" [3>3]",
                "no fragment: consume(64) []",
                "no fragment: fill_buf => \"56789\" [2>2 5>5]",
                "no fragment: consume(64) []",
                "no fragment: fill_buf => \"\" []",
                "no fragment: consume(64) []",
                "no fragment: fill_buf => \"\" []",
                "no fragment: consume(64) []",
                "no fragment: fill_buf => \"\" []",
                "no fragment: consume(64) []",
                "no fragment: finish => source 10 bytes 84d89877, target 8 bytes c954a6fe",
                "no suffix: fill_buf => \"01234567\" [8>8]",
                "no suffix: consume(64) []",
                "no suffix: fill_buf => \"AB\" [2>2]",
                "no suffix: consume(64) []",
                "no suffix: fill_buf => \"\" []",
                "no suffix: consume(64) []",
                "no suffix: fill_buf => \"\" []",
                "no suffix: consume(64) []",
                "no suffix: fill_buf => \"\" []",
                "no suffix: consume(64) []",
                "no suffix: finish => source 10 bytes 84d89877, target 10 bytes f2bde2b6",
                "insert at end: fill_buf => \"0123456789\" [10>10]",
                "insert at end: consume(64) []",
                "insert at end: fill_buf => \"AB\" []",
                "insert at end: consume(64) []",
                "insert at end: fill_buf => \"\" []",
                "insert at end: consume(64) []",
                "insert at end: fill_buf => \"\" []",
                "insert at end: consume(64) []",
                "insert at end: fill_buf => \"\" []",
                "insert at end: consume(64) []",
                "insert at end: finish => source 10 bytes 84d89877, target 12 bytes 4e3f5a1d",
                "insert at start: fill_buf => \"AB\" []",
                "insert at start: consume(64) []",
                "insert at start: fill_buf => \"0123456789\" [10>10]",
                "insert at start: consume(64) []",
                "insert at start: fill_buf => \"\" []",
                "insert at start: consume(64) []",
                "insert at start: fill_buf => \"\" []",
                "insert at start: consume(64) []",
                "insert at start: fill_buf => \"\" []",
                "insert at start: consume(64) []",
                "insert at start: finish => source 10 bytes 84d89877, target 12 bytes 2408b90c",
                "replace everything: fill_buf => \"AB\" [10>10]",
                "replace everything: consume(64) []",
                "replace everything: fill_buf => \"\" []",
                "replace everything: consume(64) []",
                "replace everything: fill_buf => \"\" []",
                "replace everything: consume(64) []",
                "replace everything: fill_buf => \"\" []",
                "replace everything: consume(64) []",
                "replace everything: fill_buf => \"\" []",
                "replace everything: consume(64) []",
                "replace everything: finish => source 10 bytes 84d89877, target 2 bytes 38164fbd",
                "remove everything: fill_buf => \"\" [10>10]",
                "remove everything: consume(64) []",
                "remove everything: fill_buf => \"\" []",
                "remove everything: consume(64) []",
                "remove everything: fill_buf => \"\" []",
                "remove everything: consume(64) []",
                "remove everything: fill_buf => \"\" []",
                "remove everything: consume(64) []",
                "remove everything: fill_buf => \"\" []",
                "remove everything: consume(64) []",
                "remove everything: finish => source 10 bytes 84d89877, target 0 bytes e3b0c442",
                "nothing to do: fill_buf => \"0123456789\" [10>10]",
                "nothing to do: consume(64) []",
                "nothing to do: fill_buf => \"\" []",
                "nothing to do: consume(64) []",
                "nothing to do: fill_buf => \"\" []",
                "nothing to do: consume(64) []",
                "nothing to do: fill_buf => \"\" []",
                "nothing to do: consume(64) []",
                "nothing to do: fill_buf => \"\" []",
                "nothing to do: consume(64) []",
                "nothing to do: finish => source 10 bytes 84d89877, target 10 bytes 84d89877",
                "empty source and fragment: fill_buf => \"\" []",
                "empty source and fragment: consume(64) []",
                "empty source and fragment: fill_buf => \"\" []",
                "empty source and fragment: consume(64) []",
                "empty source and fragment: fill_buf => \"\" []",
                "empty source and fragment: consume(64) []",
                "empty source and fragment: fill_buf => \"\" []",
                "empty source and fragment: consume(64) []",
                "empty source and fragment: fill_buf => \"\" []",
                "empty source and fragment: consume(64) []",
                "empty source and fragment: finish => source 0 bytes e3b0c442, target 0 bytes e3b0c442",
                "empty source: fill_buf => \"AB\" []",
                "empty source: consume(64) []",
                "empty source: fill_buf => \"\" []",
                "empty source: consume(64) []",
                "empty source: fill_buf => \"\" []",
                "empty source: consume(64) []",
                "empty source: fill_buf => \"\" []",
                "empty source: consume(64) []",
                "empty source: fill_buf => \"\" []",
                "empty source: consume(64) []",
                "empty source: finish => source 0 bytes e3b0c442, target 2 bytes 38164fbd",
            ],
        );
    }

    #[test]
    fn splice_chunks_long_phases() {
        let letters = |length: usize| -> Vec<u8> {
            (0..length).map(|index| b'a' + (index % 26) as u8).collect()
        };
        let data = letters(4 * COPY_CHUNK_SIZE + 5);
        let case = Case {
            expected_source_length: data.len() as u64,
            data,
            script: Vec::new(),
            offset: COPY_CHUNK_SIZE as u64 + 2,
            remove: COPY_CHUNK_SIZE as u64 + 1,
            fragment: letters(COPY_CHUNK_SIZE + 3)
                .iter()
                .map(u8::to_ascii_uppercase)
                .collect(),
        };
        let ops = [Op::FillBuf, Op::Consume(usize::MAX)].repeat(8);
        assert_lines(
            &transcript(&case, &ops),
            &[
                "fill_buf => 65536 bytes ending \"mnop\" [65536>65536]",
                "consume(18446744073709551615) []",
                "fill_buf => \"qr\" [2>2]",
                "consume(18446744073709551615) []",
                "fill_buf => 65536 bytes ending \"MNOP\" [65536>65536 1>1]",
                "consume(18446744073709551615) []",
                "fill_buf => \"QRS\" []",
                "consume(18446744073709551615) []",
                "fill_buf => 65536 bytes ending \"vwxy\" [65536>65536]",
                "consume(18446744073709551615) []",
                "fill_buf => 65536 bytes ending \"lmno\" [65536>65536]",
                "consume(18446744073709551615) []",
                "fill_buf => \"pq\" [2>2]",
                "consume(18446744073709551615) []",
                "fill_buf => \"\" []",
                "consume(18446744073709551615) []",
                "finish => source 262149 bytes 99527826, target 262151 bytes 960bbf5c",
            ],
        );
        assert_lines(
            &transcript(&case, &[Op::Read(COPY_CHUNK_SIZE + 7); 6]),
            &[
                "read(65543) => 65543 bytes ending \"BCDE\" [65536>65536 2>2 65536>65536 1>1]",
                "read(65543) => 65543 bytes ending \"opqr\" [65536>65536]",
                "read(65543) => 65543 bytes ending \"lmno\" [65536>65536]",
                "read(65543) => 65522 bytes ending \"nopq\" [2>2]",
                "read(65543) => \"\" []",
                "read(65543) => \"\" []",
                "finish => source 262149 bytes 99527826, target 262151 bytes 960bbf5c",
            ],
        );
    }
}
