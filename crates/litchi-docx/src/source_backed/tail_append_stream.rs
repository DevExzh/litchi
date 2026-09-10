//! Replayable source-backed append of caller-authored plain paragraphs.
//!
//! The event, provider, proof, and replay-store vocabulary is owned here.
//! The internal encoder owns the bounded WordprocessingML encoder and deterministic
//! cursor adapter. The fixed one-paragraph operation in
//! [`super::tail_append`] remains an independent compatibility path.

use super::{Package, tail_append};
use crate::namespace::{STRICT_WORDPROCESSINGML_NAMESPACE, WORDPROCESSINGML_NAMESPACE};
use litchi_core::{CancellationToken, ExecutionContext, Reservation, Resource};
use litchi_opc::SourceArtifactFingerprint;
use litchi_opc::source_backed::{
    SourcePartSpliceLimits, SourcePartSpliceProof, SourcePartSplicePublication,
    SourcePartSpliceReplay, SourcePartSpliceReplayError, SourcePartSpliceReplayProof,
    VerifiedDecodedReaderError,
};
use sha2::{Digest as _, Sha256};
use std::fmt;
use std::io::{self, BufRead, Read, Write};
use std::marker::PhantomData;
use std::path::Path;
use std::sync::{Arc, Mutex};
use thiserror::Error;

mod encoder;
pub mod patch;

const DEFAULT_REPLAY_WINDOW_BYTES: u64 = 64 * 1024;
const DEFAULT_MAX_AUTHORED_PARAGRAPHS: u64 = 1_000_000;
const DEFAULT_MAX_AUTHORED_EVENTS: u64 = 4_000_000;
const DEFAULT_MAX_AUTHORED_CHUNK_BYTES: u64 = 8 * 1024;
const DEFAULT_MAX_AUTHORED_TEXT_BYTES: u64 = 256 * 1024 * 1024;
const DEFAULT_MAX_AUTHORED_XML_BYTES: u64 = 256 * 1024 * 1024;
const DEFAULT_MAX_REPLAY_BYTES: u64 = 512 * 1024 * 1024;
const DEFAULT_MAX_PATCH_BYTES: u64 = 64 * 1024;
const REPLAY_HASH_CHUNK_BYTES: usize = 64 * 1024;

/// A borrowed event in a plain authored paragraph stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PlainParagraphEvent<'a> {
    /// Start one new paragraph.
    ParagraphStart,
    /// Append one borrowed UTF-8 text chunk to the open paragraph.
    TextChunk(&'a str),
    /// Close the open paragraph.
    ParagraphEnd,
}

/// A replay cursor whose text chunks remain caller-owned until the next call.
pub trait ParagraphCursor {
    /// Error reported by the caller's source.
    type Error: std::error::Error + Send + Sync + 'static;

    /// Return the next borrowed event.
    fn next<'event>(
        &'event mut self,
    ) -> std::result::Result<Option<PlainParagraphEvent<'event>>, Self::Error>;
}

/// A deterministic source that opens an equivalent event cursor for every pass.
pub trait ReplayableParagraphSource: Send + Sync {
    /// Source error type.
    type Error: std::error::Error + Send + Sync + 'static;
    /// Cursor type for one fresh source pass.
    type Cursor<'source>: ParagraphCursor<Error = Self::Error> + Send
    where
        Self: 'source;

    /// Open one fresh cursor.
    fn open<'source>(&'source self) -> std::result::Result<Self::Cursor<'source>, Self::Error>;

    /// Return a bounded provider token when durable replay is available.
    fn durable_reference(&self) -> Option<AuthoredReplayReference> {
        None
    }
}

/// A bounded caller/provider token for durable authored replay.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AuthoredReplayReference(Arc<Vec<u8>>);

impl AuthoredReplayReference {
    /// Construct a reference from bounded caller/provider token bytes.
    ///
    /// The bytes are copied once into shared immutable storage. Cloning the
    /// reference therefore shares the bounded allocation.
    pub fn try_from_bytes(
        bytes: &[u8],
        maximum: u64,
    ) -> std::result::Result<Self, AuthoredReplayError> {
        validate_replay_reference_maximum(maximum)?;
        let length = u64::try_from(bytes.len()).map_err(|_| AuthoredReplayError::Limit {
            resource: "durable replay reference bytes",
            actual: u64::MAX,
            maximum,
        })?;
        if length == 0 || length > maximum {
            return Err(AuthoredReplayError::Limit {
                resource: "durable replay reference bytes",
                actual: length,
                maximum,
            });
        }
        let mut owned = Vec::new();
        owned.try_reserve_exact(bytes.len()).map_err(|_| {
            AuthoredReplayError::Store("durable replay reference allocation failed")
        })?;
        if owned.capacity() != bytes.len() {
            return Err(AuthoredReplayError::Store(
                "durable replay reference allocation exceeded its reservation",
            ));
        }
        owned.extend_from_slice(bytes);
        Ok(Self(Arc::new(owned)))
    }

    pub(crate) fn validate_for(
        &self,
        maximum: u64,
    ) -> std::result::Result<(), AuthoredReplayError> {
        validate_replay_reference_maximum(maximum)?;
        let length = u64::try_from(self.0.len()).map_err(|_| AuthoredReplayError::Limit {
            resource: "durable replay reference bytes",
            actual: u64::MAX,
            maximum,
        })?;
        if length > maximum {
            return Err(AuthoredReplayError::Limit {
                resource: "durable replay reference bytes",
                actual: length,
                maximum,
            });
        }
        Ok(())
    }

    /// Borrow the opaque token bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.0.as_slice()
    }

    /// Return the token length.
    #[must_use]
    pub fn len(&self) -> usize {
        self.0.len()
    }

    /// Return whether the token is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}

fn validate_replay_reference_maximum(maximum: u64) -> std::result::Result<(), AuthoredReplayError> {
    if maximum == 0 || maximum == u64::MAX {
        return Err(AuthoredReplayError::Limit {
            resource: "durable replay reference bytes",
            actual: maximum,
            maximum: maximum.saturating_sub(1),
        });
    }
    if maximum > patch::ABSOLUTE_MAX_PATCH_BYTES as u64 {
        return Err(AuthoredReplayError::Limit {
            resource: "durable replay reference bytes",
            actual: maximum,
            maximum: patch::ABSOLUTE_MAX_PATCH_BYTES as u64,
        });
    }
    Ok(())
}

/// Compact authenticated facts for one complete authored stream.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct AuthoredStreamProof {
    /// Whether generated elements use Strict WordprocessingML.
    pub strict_namespace: bool,
    /// Number of completed authored paragraphs.
    pub paragraph_count: u64,
    /// Number of authored events, including start/text/end events.
    pub event_count: u64,
    /// UTF-8 bytes supplied by the caller in text events.
    pub text_bytes: u64,
    /// Exact generated XML bytes emitted by the encoder.
    pub encoded_xml_bytes: u64,
    /// Digest over unambiguous event framing and borrowed text bytes.
    pub event_sha256: [u8; 32],
    /// Digest over exact generated XML bytes.
    pub encoded_sha256: [u8; 32],
}

/// Proof returned after a replay reader reaches EOF and authenticates its pass.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct AuthoredPassProof(pub AuthoredStreamProof);

impl AuthoredPassProof {
    /// Borrow the completed pass facts.
    #[must_use]
    pub const fn proof(self) -> AuthoredStreamProof {
        self.0
    }
}

/// Error raised by an authored source, encoder, or explicit replay store.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum AuthoredReplayError {
    /// Caller source returned an error.
    #[error("authored paragraph provider failed: {0}")]
    Provider(#[source] Box<dyn std::error::Error + Send + Sync>),
    /// A stream event violated the strict event grammar.
    #[error("authored paragraph stream is invalid: {0}")]
    Invalid(&'static str),
    /// A bounded authored resource was exceeded.
    #[error("authored paragraph stream {resource} limit exceeded: {actual} > {maximum}")]
    Limit {
        /// Resource name.
        resource: &'static str,
        /// Observed value.
        actual: u64,
        /// Configured ceiling.
        maximum: u64,
    },
    /// The provider changed or emitted a different authenticated stream.
    #[error("authored paragraph replay proof changed between passes")]
    Changed,
    /// The underlying sequential replay reader failed.
    #[error("authored paragraph replay I/O failed: {0}")]
    Io(#[from] io::Error),
    /// A replay store rejected a chunk or bounded allocation.
    #[error("authored paragraph replay store refused a chunk: {0}")]
    Store(&'static str),
}

/// A replay reader with an authenticated terminal proof.
pub trait AuthoredReplayReader: Read + Send {
    /// Finish a reader that has reached EOF and return its pass proof.
    fn finish(self: Box<Self>) -> std::result::Result<AuthoredPassProof, AuthoredReplayError>;
}

/// A sealed authored byte stream that can be opened afresh for every pass.
pub trait AuthoredReplayHandle: Send + Sync {
    /// Return the immutable sealed proof.
    fn proof(&self) -> AuthoredStreamProof;
    /// Open one fresh authenticated reader.
    fn open(&self) -> std::result::Result<Box<dyn AuthoredReplayReader + '_>, AuthoredReplayError>;
    /// Open a reader bound to the current package operation.
    ///
    /// External handles may use the default implementation. Managed handles
    /// override it when their retained storage carries an older operation
    /// context, so a reopened package charges its own work and cancellation
    /// policy instead of the provider's original one. `None` values therefore
    /// deliberately clear an older managed context on a reopened unmanaged
    /// package.
    fn open_for_package<'a>(
        &'a self,
        context: Option<&'a ExecutionContext>,
        cancellation: Option<&'a CancellationToken>,
    ) -> std::result::Result<Box<dyn AuthoredReplayReader + 'a>, AuthoredReplayError> {
        let _ = (context, cancellation);
        self.open()
    }
    /// Return the provider token, if durable replay is available.
    fn durable_reference(&self) -> Option<AuthoredReplayReference>;
}

/// Explicit destination for encoded chunks produced by a one-shot producer.
pub trait AuthoredReplayStore: Send {
    /// Sealed handle type returned after all bytes have been appended.
    type Handle: AuthoredReplayHandle;

    /// Prepare caller-owned storage for one selected stream operation.
    ///
    /// The default is deliberately an external-storage policy: it performs
    /// only the operation and cancellation checks and does not allocate or
    /// claim package memory. A library-owned store may override this hook to
    /// compare its retained capacity with `limits` and reserve that capacity
    /// from `context` before the producer is invoked.
    fn prepare_for_operation(
        &mut self,
        limits: ParagraphStreamLimits,
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
    ) -> std::result::Result<(), AuthoredReplayError> {
        check_store_operation(context, cancellation)
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        let _ = limits;
        Ok(())
    }

    /// Append one bounded encoded chunk.
    fn append(&mut self, chunk: &[u8]) -> std::result::Result<(), AuthoredReplayError>;

    /// Seal the store with the authenticated authored proof.
    fn finish(
        self,
        proof: AuthoredStreamProof,
    ) -> std::result::Result<Self::Handle, AuthoredReplayError>;
}

/// Consumer of borrowed paragraph events used by one-shot producers.
pub trait ParagraphEventSink {
    /// Submit one event while its borrowed text remains valid.
    fn push<'event>(
        &mut self,
        event: PlainParagraphEvent<'event>,
    ) -> std::result::Result<(), AuthoredReplayError>;
}

/// A producer that is run exactly once into an explicit replay store.
pub trait OneShotParagraphProducer: Send {
    /// Emit all events to the supplied bounded sink.
    fn produce(
        &mut self,
        sink: &mut dyn ParagraphEventSink,
    ) -> std::result::Result<(), AuthoredReplayError>;
}

impl<F> OneShotParagraphProducer for F
where
    F: FnMut(&mut dyn ParagraphEventSink) -> std::result::Result<(), AuthoredReplayError> + Send,
{
    fn produce(
        &mut self,
        sink: &mut dyn ParagraphEventSink,
    ) -> std::result::Result<(), AuthoredReplayError> {
        self(sink)
    }
}

/// Finite bounds for a multi-paragraph source-backed append.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParagraphStreamLimits {
    /// Existing source-backed DOCX limits.
    pub source: tail_append::Limits,
    /// Maximum completed authored paragraphs.
    pub max_authored_paragraphs: u64,
    /// Maximum authored events.
    pub max_authored_events: u64,
    /// Maximum UTF-8 bytes in one borrowed text chunk.
    pub max_authored_chunk_bytes: u64,
    /// Maximum total UTF-8 input bytes.
    pub max_authored_text_bytes: u64,
    /// Maximum complete generated XML bytes.
    pub max_authored_xml_bytes: u64,
    /// Maximum explicit replay-store retention.
    pub max_replay_bytes: u64,
    /// Maximum one-reader replay window.
    pub max_replay_window_bytes: u64,
    /// Maximum compact durable patch bytes.
    pub max_patch_bytes: u64,
}

impl Default for ParagraphStreamLimits {
    fn default() -> Self {
        Self {
            source: tail_append::Limits::default(),
            max_authored_paragraphs: DEFAULT_MAX_AUTHORED_PARAGRAPHS,
            max_authored_events: DEFAULT_MAX_AUTHORED_EVENTS,
            max_authored_chunk_bytes: DEFAULT_MAX_AUTHORED_CHUNK_BYTES,
            max_authored_text_bytes: DEFAULT_MAX_AUTHORED_TEXT_BYTES,
            max_authored_xml_bytes: DEFAULT_MAX_AUTHORED_XML_BYTES,
            max_replay_bytes: DEFAULT_MAX_REPLAY_BYTES,
            max_replay_window_bytes: DEFAULT_REPLAY_WINDOW_BYTES,
            max_patch_bytes: DEFAULT_MAX_PATCH_BYTES,
        }
    }
}

impl ParagraphStreamLimits {
    /// Construct a policy with existing source limits and authored ceilings.
    #[must_use]
    pub const fn new(source: tail_append::Limits) -> Self {
        Self {
            source,
            max_authored_paragraphs: DEFAULT_MAX_AUTHORED_PARAGRAPHS,
            max_authored_events: DEFAULT_MAX_AUTHORED_EVENTS,
            max_authored_chunk_bytes: DEFAULT_MAX_AUTHORED_CHUNK_BYTES,
            max_authored_text_bytes: DEFAULT_MAX_AUTHORED_TEXT_BYTES,
            max_authored_xml_bytes: DEFAULT_MAX_AUTHORED_XML_BYTES,
            max_replay_bytes: DEFAULT_MAX_REPLAY_BYTES,
            max_replay_window_bytes: DEFAULT_REPLAY_WINDOW_BYTES,
            max_patch_bytes: DEFAULT_MAX_PATCH_BYTES,
        }
    }

    /// Reconstruct the stream policy retained by a durable forward patch.
    pub fn from_patch_limits(value: patch::PatchLimits) -> Result<Self, Error> {
        value.validate().map_err(Error::Patch)?;
        Ok(Self {
            source: tail_append::Limits::new(
                value.max_source_xml_bytes,
                value.max_text_bytes,
                value.max_fragment_bytes,
                value.max_candidate_xml_bytes,
                value.max_events,
                value.max_depth,
                value.max_paragraphs,
                value.max_settings_xml_bytes,
                value.max_workspace_bytes,
                value.max_output_bytes,
                value.max_token_bytes,
            ),
            max_authored_paragraphs: value.max_authored_paragraphs,
            max_authored_events: value.max_authored_events,
            max_authored_chunk_bytes: value.max_authored_chunk_bytes,
            max_authored_text_bytes: value.max_authored_text_bytes,
            max_authored_xml_bytes: value.max_authored_xml_bytes,
            max_replay_bytes: value.max_replay_bytes,
            max_replay_window_bytes: value.max_replay_window_bytes,
            max_patch_bytes: value.max_patch_bytes,
        })
    }

    /// Validate all independent finite ceilings before consuming authored input.
    pub fn validate(self) -> Result<(), Error> {
        self.source.validate_for_stream()?;
        for (resource, value) in [
            ("authored paragraphs", self.max_authored_paragraphs),
            ("authored events", self.max_authored_events),
            ("authored chunk bytes", self.max_authored_chunk_bytes),
            ("authored text bytes", self.max_authored_text_bytes),
            ("authored XML bytes", self.max_authored_xml_bytes),
            ("replay bytes", self.max_replay_bytes),
            ("replay window bytes", self.max_replay_window_bytes),
            ("patch bytes", self.max_patch_bytes),
        ] {
            if value == 0 || value == u64::MAX {
                return Err(Error::Limit {
                    resource,
                    actual: value,
                    maximum: value.saturating_sub(1),
                });
            }
        }
        let chunk_window = self
            .max_authored_chunk_bytes
            .checked_mul(6)
            .ok_or(Error::Limit {
                resource: "replay window bytes",
                actual: u64::MAX,
                maximum: self.max_replay_window_bytes,
            })?;
        if chunk_window > self.max_replay_window_bytes {
            return Err(Error::Limit {
                resource: "replay window bytes",
                actual: chunk_window,
                maximum: self.max_replay_window_bytes,
            });
        }
        if self.max_patch_bytes > patch::ABSOLUTE_MAX_PATCH_BYTES as u64 {
            return Err(Error::Limit {
                resource: "patch bytes",
                actual: self.max_patch_bytes,
                maximum: patch::ABSOLUTE_MAX_PATCH_BYTES as u64,
            });
        }
        Ok(())
    }
}

/// Error returned by the stream transaction.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// Existing DOCX source/topology/scanner failure.
    #[error(transparent)]
    TailAppend(#[from] tail_append::Error),
    /// Authored provider or store failure.
    #[error(transparent)]
    Replay(#[from] AuthoredReplayError),
    /// OPC physical preparation/publication failure.
    #[error(transparent)]
    Opc(#[from] litchi_opc::error::OpcError),
    /// Durable forward-patch proof or provider failure.
    #[error(transparent)]
    Patch(#[from] patch::PatchError),
    /// An independent stream limit was exceeded.
    #[error("paragraph stream {resource} limit exceeded: {actual} > {maximum}")]
    Limit {
        /// Resource name.
        resource: &'static str,
        /// Observed value.
        actual: u64,
        /// Configured ceiling.
        maximum: u64,
    },
    /// A stream state or proof invariant failed.
    #[error("paragraph stream validation failed: {0}")]
    Invalid(&'static str),
}

/// Result type for the replayable paragraph stream operation.
pub type Result<T, E = Error> = std::result::Result<T, E>;

fn check_limit(resource: &'static str, actual: u64, maximum: u64) -> Result<(), Error> {
    if actual > maximum {
        return Err(Error::Limit {
            resource,
            actual,
            maximum,
        });
    }
    Ok(())
}

fn map_replay_io(error: AuthoredReplayError) -> io::Error {
    io::Error::other(error)
}

fn check_store_operation(
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
) -> std::result::Result<(), litchi_core::ExecutionError> {
    if let Some(context) = context {
        context.check()?;
    }
    if let Some(cancellation) = cancellation {
        cancellation.check()?;
    }
    Ok(())
}

fn hash_replay_chunks(
    bytes: &[u8],
    hash: &mut Sha256,
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
) -> std::result::Result<(), AuthoredReplayError> {
    for chunk in bytes.chunks(REPLAY_HASH_CHUNK_BYTES) {
        check_store_operation(context, cancellation)
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        if let Some(context) = context {
            context
                .consume(Resource::Work, chunk.len() as u64)
                .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        }
        hash.update(chunk);
    }
    Ok(())
}

/// A bounded in-memory explicit replay store.
#[derive(Debug)]
pub struct MemoryReplayStore {
    bytes: Vec<u8>,
    maximum: u64,
    durable_reference: Option<AuthoredReplayReference>,
    memory_reservation: Option<Reservation>,
    object_reservation: Option<Reservation>,
    operation_context: Option<ExecutionContext>,
    operation_cancellation: Option<CancellationToken>,
}

#[derive(Debug)]
struct MemoryReplayStorage {
    bytes: Vec<u8>,
    _memory_reservation: Option<Reservation>,
    _object_reservation: Option<Reservation>,
    operation_context: Option<ExecutionContext>,
    operation_cancellation: Option<CancellationToken>,
}

impl MemoryReplayStore {
    /// Create an empty store with an explicit retained-byte ceiling.
    pub fn new(maximum: u64) -> Result<Self, AuthoredReplayError> {
        if maximum == 0 || maximum == u64::MAX {
            return Err(AuthoredReplayError::Limit {
                resource: "replay bytes",
                actual: maximum,
                maximum: maximum.saturating_sub(1),
            });
        }
        usize::try_from(maximum).map_err(|_| AuthoredReplayError::Limit {
            resource: "replay bytes",
            actual: maximum,
            maximum: usize::MAX as u64,
        })?;
        Ok(Self {
            bytes: Vec::new(),
            maximum,
            durable_reference: None,
            memory_reservation: None,
            object_reservation: None,
            operation_context: None,
            operation_cancellation: None,
        })
    }

    /// Attach an explicit provider reference to the resulting sealed handle.
    #[must_use]
    pub fn with_durable_reference(mut self, reference: AuthoredReplayReference) -> Self {
        self.durable_reference = Some(reference);
        self
    }

    fn prepare_storage(
        &mut self,
        maximum: u64,
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
    ) -> Result<(), AuthoredReplayError> {
        check_store_operation(context, cancellation)
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        let capacity = usize::try_from(maximum).map_err(|_| AuthoredReplayError::Limit {
            resource: "replay bytes",
            actual: maximum,
            maximum: usize::MAX as u64,
        })?;
        if self.memory_reservation.is_none() {
            let reservation = context
                .map(|context| {
                    context
                        .reserve(Resource::Memory, maximum)
                        .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))
                })
                .transpose()?;
            self.memory_reservation = reservation;
        }
        if self.bytes.capacity() > capacity {
            return Err(AuthoredReplayError::Store(
                "replay store capacity exceeds its reservation",
            ));
        }
        if self.bytes.capacity() < capacity {
            self.bytes
                .try_reserve_exact(capacity - self.bytes.len())
                .map_err(|_| AuthoredReplayError::Store("replay store allocation failed"))?;
            if self.bytes.capacity() != capacity {
                return Err(AuthoredReplayError::Store(
                    "replay store allocation exceeded its reservation",
                ));
            }
        }
        Ok(())
    }
}

/// Sealed handle returned by [`MemoryReplayStore`].
#[derive(Debug, Clone)]
pub struct MemoryReplayHandle {
    storage: Arc<MemoryReplayStorage>,
    proof: AuthoredStreamProof,
    durable_reference: Option<AuthoredReplayReference>,
}

struct MemoryReplayReader<'a> {
    // A reopened package supplies these instead of the storage owner's
    // original context. The outer bound reader charges Work in that case.
    operation_context: Option<&'a ExecutionContext>,
    operation_cancellation: Option<&'a CancellationToken>,
    charge_work: bool,
    _object_reservation: Option<Reservation>,
    inner: io::Cursor<&'a [u8]>,
    proof: AuthoredStreamProof,
    bytes: u64,
    hash: Sha256,
    failed: bool,
}

impl Read for MemoryReplayReader<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.failed {
            return Err(io::Error::other(
                "authored replay reader is poisoned after a terminal error",
            ));
        }
        check_store_operation(self.operation_context, self.operation_cancellation).map_err(
            |error| {
                self.failed = true;
                map_replay_io(AuthoredReplayError::Provider(Box::new(error)))
            },
        )?;
        let count = self.inner.read(output)?;
        if count != 0 {
            for chunk in output[..count].chunks(REPLAY_HASH_CHUNK_BYTES) {
                check_store_operation(self.operation_context, self.operation_cancellation)
                    .map_err(|error| {
                        self.failed = true;
                        map_replay_io(AuthoredReplayError::Provider(Box::new(error)))
                    })?;
                if self.charge_work {
                    let Some(context) = self.operation_context else {
                        self.hash.update(chunk);
                        continue;
                    };
                    context
                        .consume(Resource::Work, chunk.len() as u64)
                        .map_err(|error| {
                            self.failed = true;
                            map_replay_io(AuthoredReplayError::Provider(Box::new(error)))
                        })?;
                }
                self.hash.update(chunk);
            }
            self.bytes = self.bytes.checked_add(count as u64).ok_or_else(|| {
                self.failed = true;
                io::Error::other("authored replay length overflow")
            })?;
        }
        Ok(count)
    }
}

impl AuthoredReplayReader for MemoryReplayReader<'_> {
    fn finish(self: Box<Self>) -> Result<AuthoredPassProof, AuthoredReplayError> {
        if self.failed {
            return Err(AuthoredReplayError::Io(io::Error::other(
                "authored replay reader is poisoned after a terminal error",
            )));
        }
        check_store_operation(self.operation_context, self.operation_cancellation)
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        let actual_hash: [u8; 32] = self.hash.finalize().into();
        let actual = AuthoredStreamProof {
            encoded_xml_bytes: self.bytes,
            encoded_sha256: actual_hash,
            ..self.proof
        };
        if actual != self.proof {
            return Err(AuthoredReplayError::Changed);
        }
        Ok(AuthoredPassProof(actual))
    }
}

impl AuthoredReplayHandle for MemoryReplayHandle {
    fn proof(&self) -> AuthoredStreamProof {
        self.proof
    }

    fn open(&self) -> Result<Box<dyn AuthoredReplayReader + '_>, AuthoredReplayError> {
        self.open_reader(
            self.storage.operation_context.as_ref(),
            self.storage.operation_cancellation.as_ref(),
            true,
        )
    }

    fn open_for_package<'a>(
        &'a self,
        context: Option<&'a ExecutionContext>,
        cancellation: Option<&'a CancellationToken>,
    ) -> Result<Box<dyn AuthoredReplayReader + 'a>, AuthoredReplayError> {
        self.open_reader(context, cancellation, false)
    }

    fn durable_reference(&self) -> Option<AuthoredReplayReference> {
        self.durable_reference.clone()
    }
}

impl MemoryReplayHandle {
    fn open_reader<'a>(
        &'a self,
        operation_context: Option<&'a ExecutionContext>,
        operation_cancellation: Option<&'a CancellationToken>,
        charge_work: bool,
    ) -> Result<Box<dyn AuthoredReplayReader + 'a>, AuthoredReplayError> {
        check_store_operation(operation_context, operation_cancellation)
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        let object_reservation = operation_context
            .map(|context| {
                context
                    .reserve(Resource::Objects, 1)
                    .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))
            })
            .transpose()?;
        Ok(Box::new(MemoryReplayReader {
            operation_context,
            operation_cancellation,
            charge_work,
            _object_reservation: object_reservation,
            inner: io::Cursor::new(self.storage.bytes.as_slice()),
            proof: self.proof,
            bytes: 0,
            hash: Sha256::new(),
            failed: false,
        }))
    }
}

impl SourcePartSpliceReplay for MemoryReplayHandle {
    fn proof(&self) -> SourcePartSpliceReplayProof {
        SourcePartSpliceReplayProof {
            encoded_len: self.proof.encoded_xml_bytes,
            encoded_sha256: self.proof.encoded_sha256,
        }
    }

    fn open(&self) -> Result<Box<dyn Read + '_>, SourcePartSpliceReplayError> {
        let reader = <Self as AuthoredReplayHandle>::open(self)
            .map_err(|error| SourcePartSpliceReplayError::Provider(Box::new(error)))?;
        Ok(Box::new(AuthenticatedReader {
            inner: Some(reader),
            identity: None,
            durable_reference: None,
            expected: self.proof,
            bytes: 0,
            hash: Sha256::new(),
            finished: false,
            poisoned: false,
        }))
    }
}

impl AuthoredReplayStore for MemoryReplayStore {
    type Handle = MemoryReplayHandle;

    fn prepare_for_operation(
        &mut self,
        limits: ParagraphStreamLimits,
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
    ) -> Result<(), AuthoredReplayError> {
        if self.maximum > limits.max_replay_bytes {
            return Err(AuthoredReplayError::Limit {
                resource: "replay bytes",
                actual: self.maximum,
                maximum: limits.max_replay_bytes,
            });
        }
        if let Some(reference) = self.durable_reference.as_ref() {
            reference.validate_for(limits.max_patch_bytes)?;
        }
        self.operation_context = context.cloned();
        self.operation_cancellation = cancellation.cloned();
        if self.object_reservation.is_none() {
            self.object_reservation = context
                .map(|context| {
                    context
                        .reserve(Resource::Objects, 1)
                        .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))
                })
                .transpose()?;
        }
        self.prepare_storage(self.maximum, context, cancellation)
    }

    fn append(&mut self, chunk: &[u8]) -> Result<(), AuthoredReplayError> {
        let current = u64::try_from(self.bytes.len()).map_err(|_| AuthoredReplayError::Limit {
            resource: "replay bytes",
            actual: u64::MAX,
            maximum: self.maximum,
        })?;
        let amount = u64::try_from(chunk.len()).map_err(|_| AuthoredReplayError::Limit {
            resource: "replay bytes",
            actual: u64::MAX,
            maximum: self.maximum,
        })?;
        let next = current
            .checked_add(amount)
            .ok_or(AuthoredReplayError::Limit {
                resource: "replay bytes",
                actual: u64::MAX,
                maximum: self.maximum,
            })?;
        if next > self.maximum {
            return Err(AuthoredReplayError::Limit {
                resource: "replay bytes",
                actual: next,
                maximum: self.maximum,
            });
        }
        self.prepare_storage(self.maximum, None, None)?;
        self.bytes.extend_from_slice(chunk);
        Ok(())
    }

    fn finish(mut self, proof: AuthoredStreamProof) -> Result<Self::Handle, AuthoredReplayError> {
        let length = u64::try_from(self.bytes.len()).map_err(|_| AuthoredReplayError::Limit {
            resource: "replay bytes",
            actual: u64::MAX,
            maximum: self.maximum,
        })?;
        check_store_operation(
            self.operation_context.as_ref(),
            self.operation_cancellation.as_ref(),
        )
        .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        let mut digest = Sha256::new();
        hash_replay_chunks(
            &self.bytes,
            &mut digest,
            self.operation_context.as_ref(),
            self.operation_cancellation.as_ref(),
        )?;
        let hash: [u8; 32] = digest.finalize().into();
        if length != proof.encoded_xml_bytes || hash != proof.encoded_sha256 {
            return Err(AuthoredReplayError::Changed);
        }
        let storage = Arc::new(MemoryReplayStorage {
            bytes: self.bytes,
            _memory_reservation: self.memory_reservation.take(),
            _object_reservation: self.object_reservation.take(),
            operation_context: self.operation_context.take(),
            operation_cancellation: self.operation_cancellation.take(),
        });
        Ok(MemoryReplayHandle {
            storage,
            proof,
            durable_reference: self.durable_reference,
        })
    }
}

/// Adapt an authenticated DOCX replay handle to OPC's byte-only seam.
struct OpcReplayAdapter {
    handle: Arc<dyn AuthoredReplayHandle>,
    proof: AuthoredStreamProof,
    durable_reference: Option<AuthoredReplayReference>,
}

impl OpcReplayAdapter {
    fn new(handle: Arc<dyn AuthoredReplayHandle>) -> Self {
        let proof = handle.proof();
        let durable_reference = handle.durable_reference();
        Self {
            handle,
            proof,
            durable_reference,
        }
    }
}

struct AuthenticatedReader<'a> {
    inner: Option<Box<dyn AuthoredReplayReader + 'a>>,
    identity: Option<Arc<dyn AuthoredReplayHandle>>,
    durable_reference: Option<AuthoredReplayReference>,
    expected: AuthoredStreamProof,
    bytes: u64,
    hash: Sha256,
    finished: bool,
    poisoned: bool,
}

impl AuthenticatedReader<'_> {
    fn finish_inner(&mut self) -> io::Result<()> {
        if self.finished {
            return Ok(());
        }
        if self.poisoned {
            return Err(io::Error::other(
                "authored replay reader is poisoned after a terminal error",
            ));
        }
        let inner = match self.inner.take() {
            Some(inner) => inner,
            None => {
                self.poisoned = true;
                return Err(io::Error::other(
                    "authored replay reader was already finished",
                ));
            },
        };
        let pass = match inner.finish() {
            Ok(pass) => pass.proof(),
            Err(error) => {
                self.poisoned = true;
                return Err(map_replay_io(error));
            },
        };
        if let Some(identity) = self.identity.as_ref() {
            if identity.proof() != self.expected
                || identity.durable_reference() != self.durable_reference
            {
                self.poisoned = true;
                return Err(map_replay_io(AuthoredReplayError::Changed));
            }
        }
        let encoded_sha256: [u8; 32] = self.hash.clone().finalize().into();
        let actual = AuthoredStreamProof {
            encoded_xml_bytes: self.bytes,
            encoded_sha256,
            ..pass
        };
        if actual != self.expected || pass.encoded_xml_bytes != self.bytes {
            self.poisoned = true;
            return Err(map_replay_io(AuthoredReplayError::Changed));
        }
        self.finished = true;
        Ok(())
    }
}

impl Read for AuthenticatedReader<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.poisoned {
            return Err(io::Error::other(
                "authored replay reader is poisoned after a terminal error",
            ));
        }
        if output.is_empty() {
            return Ok(0);
        }
        if self.finished {
            return Ok(0);
        }
        if let Some(identity) = self.identity.as_ref() {
            if identity.proof() != self.expected
                || identity.durable_reference() != self.durable_reference
            {
                self.poisoned = true;
                return Err(map_replay_io(AuthoredReplayError::Changed));
            }
        }
        let inner = self
            .inner
            .as_mut()
            .ok_or_else(|| io::Error::other("authored replay reader is unavailable"))?;
        let count = match inner.read(output) {
            Ok(count) => count,
            Err(error) => {
                self.poisoned = true;
                return Err(error);
            },
        };
        if count > output.len() {
            self.poisoned = true;
            return Err(map_replay_io(AuthoredReplayError::Invalid(
                "authored replay reader returned more bytes than requested",
            )));
        }
        if count == 0 {
            self.finish_inner()?;
        } else {
            let count_u64 = u64::try_from(count).map_err(|_| {
                self.poisoned = true;
                map_replay_io(AuthoredReplayError::Invalid(
                    "authored replay byte count does not fit in u64",
                ))
            })?;
            let next = match self.bytes.checked_add(count_u64) {
                Some(bytes) => bytes,
                None => {
                    self.poisoned = true;
                    return Err(io::Error::other("authored replay length overflow"));
                },
            };
            if next > self.expected.encoded_xml_bytes {
                self.poisoned = true;
                return Err(map_replay_io(AuthoredReplayError::Changed));
            }
            self.bytes = next;
            self.hash.update(&output[..count]);
        }
        Ok(count)
    }
}

impl SourcePartSpliceReplay for OpcReplayAdapter {
    fn proof(&self) -> SourcePartSpliceReplayProof {
        SourcePartSpliceReplayProof {
            encoded_len: self.proof.encoded_xml_bytes,
            encoded_sha256: self.proof.encoded_sha256,
        }
    }

    fn open(&self) -> Result<Box<dyn Read + '_>, SourcePartSpliceReplayError> {
        if self.handle.proof() != self.proof
            || self.handle.durable_reference() != self.durable_reference
        {
            return Err(SourcePartSpliceReplayError::Provider(Box::new(
                AuthoredReplayError::Changed,
            )));
        }
        let reader = self
            .handle
            .open()
            .map_err(|error| SourcePartSpliceReplayError::Provider(Box::new(error)))?;
        Ok(Box::new(AuthenticatedReader {
            inner: Some(reader),
            identity: Some(Arc::clone(&self.handle)),
            durable_reference: self.durable_reference.clone(),
            expected: self.proof,
            bytes: 0,
            hash: Sha256::new(),
            finished: false,
            poisoned: false,
        }))
    }
}

struct BoundReplayHandle {
    inner: Arc<dyn AuthoredReplayHandle>,
    proof: AuthoredStreamProof,
    durable_reference: Option<AuthoredReplayReference>,
    access_context: Option<ExecutionContext>,
    access_cancellation: Option<CancellationToken>,
    replace_access_policy: bool,
}

impl BoundReplayHandle {
    fn bind(
        inner: Arc<dyn AuthoredReplayHandle>,
        proof: AuthoredStreamProof,
        durable_reference: Option<AuthoredReplayReference>,
        access_context: Option<ExecutionContext>,
        access_cancellation: Option<CancellationToken>,
        replace_access_policy: bool,
    ) -> Result<Arc<dyn AuthoredReplayHandle>, AuthoredReplayError> {
        if inner.proof() != proof || inner.durable_reference() != durable_reference {
            return Err(AuthoredReplayError::Changed);
        }
        Ok(Arc::new(Self {
            inner,
            proof,
            durable_reference,
            access_context,
            access_cancellation,
            replace_access_policy,
        }))
    }

    fn verify_inner(&self) -> std::result::Result<(), AuthoredReplayError> {
        if self.inner.proof() != self.proof
            || self.inner.durable_reference() != self.durable_reference
        {
            return Err(AuthoredReplayError::Changed);
        }
        Ok(())
    }
}

struct BoundReplayReader<'a> {
    inner: Option<Box<dyn AuthoredReplayReader + 'a>>,
    owner: Arc<dyn AuthoredReplayHandle>,
    expected: AuthoredStreamProof,
    durable_reference: Option<AuthoredReplayReference>,
    // The bound reader owns per-pass package checks and Work charges. Managed
    // memory readers suppress their duplicate Work charge when this is set.
    access_context: Option<ExecutionContext>,
    access_cancellation: Option<CancellationToken>,
    poisoned: bool,
}

impl Read for BoundReplayReader<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.poisoned {
            return Err(io::Error::other(
                "authored replay reader is poisoned after a terminal error",
            ));
        }
        check_store_operation(
            self.access_context.as_ref(),
            self.access_cancellation.as_ref(),
        )
        .map_err(|error| {
            self.poisoned = true;
            map_replay_io(AuthoredReplayError::Provider(Box::new(error)))
        })?;
        let inner = self
            .inner
            .as_mut()
            .ok_or_else(|| io::Error::other("authored replay reader was already finished"))?;
        let count = match inner.read(output) {
            Ok(count) => count,
            Err(error) => {
                self.poisoned = true;
                return Err(error);
            },
        };
        if count > output.len() {
            self.poisoned = true;
            return Err(map_replay_io(AuthoredReplayError::Invalid(
                "authored replay reader returned more bytes than requested",
            )));
        }
        for chunk in output[..count].chunks(REPLAY_HASH_CHUNK_BYTES) {
            check_store_operation(
                self.access_context.as_ref(),
                self.access_cancellation.as_ref(),
            )
            .map_err(|error| {
                self.poisoned = true;
                map_replay_io(AuthoredReplayError::Provider(Box::new(error)))
            })?;
            if let Some(context) = self.access_context.as_ref() {
                context
                    .consume(Resource::Work, chunk.len() as u64)
                    .map_err(|error| {
                        self.poisoned = true;
                        map_replay_io(AuthoredReplayError::Provider(Box::new(error)))
                    })?;
            }
        }
        check_store_operation(
            self.access_context.as_ref(),
            self.access_cancellation.as_ref(),
        )
        .map_err(|error| {
            self.poisoned = true;
            map_replay_io(AuthoredReplayError::Provider(Box::new(error)))
        })?;
        Ok(count)
    }
}

impl AuthoredReplayReader for BoundReplayReader<'_> {
    fn finish(self: Box<Self>) -> Result<AuthoredPassProof, AuthoredReplayError> {
        let BoundReplayReader {
            mut inner,
            owner,
            expected,
            durable_reference,
            access_context,
            access_cancellation,
            poisoned,
        } = *self;
        if poisoned {
            return Err(AuthoredReplayError::Io(io::Error::other(
                "authored replay reader is poisoned after a terminal error",
            )));
        }
        let reader = inner.take().ok_or(AuthoredReplayError::Invalid(
            "authored replay reader was already finished",
        ))?;
        check_store_operation(access_context.as_ref(), access_cancellation.as_ref())
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        let pass = reader.finish()?;
        check_store_operation(access_context.as_ref(), access_cancellation.as_ref())
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        if pass.proof() != expected
            || owner.proof() != expected
            || owner.durable_reference() != durable_reference
        {
            return Err(AuthoredReplayError::Changed);
        }
        Ok(pass)
    }
}

fn bind_replay_handle(
    inner: Arc<dyn AuthoredReplayHandle>,
    proof: AuthoredStreamProof,
    limits: ParagraphStreamLimits,
) -> Result<Arc<dyn AuthoredReplayHandle>, AuthoredReplayError> {
    bind_replay_handle_with_context(inner, proof, limits, None, None, false)
}

fn bind_replay_handle_with_context(
    inner: Arc<dyn AuthoredReplayHandle>,
    proof: AuthoredStreamProof,
    limits: ParagraphStreamLimits,
    access_context: Option<&ExecutionContext>,
    access_cancellation: Option<&CancellationToken>,
    replace_access_policy: bool,
) -> Result<Arc<dyn AuthoredReplayHandle>, AuthoredReplayError> {
    check_store_operation(access_context, access_cancellation)
        .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
    let durable_reference = inner.durable_reference();
    if let Some(reference) = durable_reference.as_ref() {
        reference.validate_for(limits.max_patch_bytes)?;
    }
    BoundReplayHandle::bind(
        inner,
        proof,
        durable_reference,
        access_context.cloned(),
        access_cancellation.cloned(),
        replace_access_policy,
    )
}

impl AuthoredReplayHandle for BoundReplayHandle {
    fn proof(&self) -> AuthoredStreamProof {
        self.proof
    }

    fn open(&self) -> Result<Box<dyn AuthoredReplayReader + '_>, AuthoredReplayError> {
        self.verify_inner()?;
        check_store_operation(
            self.access_context.as_ref(),
            self.access_cancellation.as_ref(),
        )
        .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        let reader = if self.replace_access_policy {
            self.inner.open_for_package(
                self.access_context.as_ref(),
                self.access_cancellation.as_ref(),
            )?
        } else {
            self.inner.open()?
        };
        check_store_operation(
            self.access_context.as_ref(),
            self.access_cancellation.as_ref(),
        )
        .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        Ok(Box::new(BoundReplayReader {
            inner: Some(reader),
            owner: Arc::clone(&self.inner),
            expected: self.proof,
            durable_reference: self.durable_reference.clone(),
            access_context: self.access_context.clone(),
            access_cancellation: self.access_cancellation.clone(),
            poisoned: false,
        }))
    }

    fn durable_reference(&self) -> Option<AuthoredReplayReference> {
        self.durable_reference.clone()
    }
}

trait AuthoredInput: Send + Sync {
    fn seal(
        &self,
        strict_namespace: bool,
        limits: ParagraphStreamLimits,
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
    ) -> Result<Arc<dyn AuthoredReplayHandle>, Error>;
}

struct DeterministicInput<S> {
    source: Arc<S>,
}

impl<S> AuthoredInput for DeterministicInput<S>
where
    S: ReplayableParagraphSource + 'static,
{
    fn seal(
        &self,
        strict_namespace: bool,
        limits: ParagraphStreamLimits,
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
    ) -> Result<Arc<dyn AuthoredReplayHandle>, Error> {
        let handle = encoder::DeterministicReplayHandle::seal(
            Arc::clone(&self.source),
            context,
            cancellation,
            limits,
            strict_namespace,
            encoder::CursorAccounting::default(),
        )?;
        let handle: Arc<dyn AuthoredReplayHandle> = Arc::new(handle);
        let proof = handle.proof();
        bind_replay_handle(handle, proof, limits).map_err(Error::Replay)
    }
}

struct ProducerInput<P, R> {
    producer_and_store: Mutex<Option<(P, R)>>,
}

impl<P, R> AuthoredInput for ProducerInput<P, R>
where
    P: OneShotParagraphProducer + 'static,
    R: AuthoredReplayStore + 'static,
    R::Handle: 'static,
{
    fn seal(
        &self,
        strict_namespace: bool,
        limits: ParagraphStreamLimits,
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
    ) -> Result<Arc<dyn AuthoredReplayHandle>, Error> {
        let (mut producer, mut store) = self
            .producer_and_store
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
            .ok_or(Error::Invalid("one-shot producer was already consumed"))?;
        store.prepare_for_operation(limits, context, cancellation)?;
        let chunk_sink = StoreChunkSink { store: &mut store };
        let mut encoder = encoder::StreamingParagraphEncoder::new(
            context,
            cancellation,
            limits,
            strict_namespace,
            encoder::CursorAccounting::default(),
            chunk_sink,
        )?;
        let mut event_sink = EncoderEventSink {
            encoder: &mut encoder,
        };
        producer.produce(&mut event_sink)?;
        let proof = encoder.finish()?.proof();
        let handle = store.finish(proof)?;
        let handle: Arc<dyn AuthoredReplayHandle> = Arc::new(handle);
        if handle.proof() != proof {
            return Err(Error::Replay(AuthoredReplayError::Changed));
        }
        bind_replay_handle(handle, proof, limits).map_err(Error::Replay)
    }
}

struct StoreChunkSink<'a, R: AuthoredReplayStore> {
    store: &'a mut R,
}

impl<R: AuthoredReplayStore> encoder::ChunkSink for StoreChunkSink<'_, R> {
    fn accept(&mut self, chunk: &[u8]) -> Result<(), AuthoredReplayError> {
        self.store.append(chunk)
    }
}

struct EncoderEventSink<'a, S: encoder::ChunkSink> {
    encoder: &'a mut encoder::StreamingParagraphEncoder<S>,
}

impl<S: encoder::ChunkSink> ParagraphEventSink for EncoderEventSink<'_, S> {
    fn push<'event>(
        &mut self,
        event: PlainParagraphEvent<'event>,
    ) -> Result<(), AuthoredReplayError> {
        self.encoder.push_event(event)
    }
}

#[derive(Clone)]
struct PatchExpectations {
    source: patch::StreamSourceProof,
    authored: AuthoredStreamProof,
    candidate: patch::StreamCandidateProof,
    replay_reference: Option<AuthoredReplayReference>,
}

impl PatchExpectations {
    fn from_patch(value: &patch::ParagraphStreamPatch) -> Self {
        Self {
            source: value.source().clone(),
            authored: value.authored_proof(),
            candidate: value.candidate(),
            replay_reference: value.replay_reference().cloned(),
        }
    }
}

struct ResolvedInput {
    handle: Arc<dyn AuthoredReplayHandle>,
    expected: AuthoredStreamProof,
    expected_reference: AuthoredReplayReference,
}

impl AuthoredInput for ResolvedInput {
    fn seal(
        &self,
        strict_namespace: bool,
        limits: ParagraphStreamLimits,
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
    ) -> Result<Arc<dyn AuthoredReplayHandle>, Error> {
        check_store_operation(context, cancellation)
            .map_err(|error| Error::Replay(AuthoredReplayError::Provider(Box::new(error))))?;
        let proof = self.handle.proof();
        if proof != self.expected
            || proof.strict_namespace != strict_namespace
            || self.handle.durable_reference().as_ref() != Some(&self.expected_reference)
        {
            return Err(Error::Patch(patch::PatchError::ReplayProofMismatch));
        }
        let bound = bind_replay_handle_with_context(
            Arc::clone(&self.handle),
            self.expected,
            limits,
            context,
            cancellation,
            true,
        )
        .map_err(Error::Replay)?;
        check_store_operation(context, cancellation)
            .map_err(|error| Error::Replay(AuthoredReplayError::Provider(Box::new(error))))?;
        Ok(bound)
    }
}

/// An edit consuming either a deterministic event source or a one-shot producer.
pub struct ParagraphStreamEdit<'package, S = ()> {
    package: &'package Package,
    input: Arc<dyn AuthoredInput>,
    limits: ParagraphStreamLimits,
    cancellation: Option<CancellationToken>,
    expected: Option<PatchExpectations>,
    marker: PhantomData<fn() -> S>,
}

impl<S> fmt::Debug for ParagraphStreamEdit<'_, S> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ParagraphStreamEdit")
            .field("limits", &self.limits)
            .finish_non_exhaustive()
    }
}

impl<'package, S> ParagraphStreamEdit<'package, S> {
    /// Replace the finite stream policy before preparation.
    #[must_use]
    pub fn with_limits(mut self, limits: ParagraphStreamLimits) -> Self {
        self.limits = limits;
        self
    }

    /// Attach an additional cooperative cancellation token.
    #[must_use]
    pub fn with_cancellation_token(mut self, token: &CancellationToken) -> Self {
        self.cancellation = Some(token.clone());
        self
    }

    /// Return the selected finite stream policy.
    #[must_use]
    pub const fn limits(&self) -> ParagraphStreamLimits {
        self.limits
    }

    /// Prepare source, authored, candidate, and OPC proofs without touching a sink.
    pub fn prepare(self) -> Result<ParagraphStreamPlan<'package>> {
        prepare_stream(self)
    }

    /// Prepare and consume the stream edit as a named commit product.
    pub fn commit(self) -> Result<ParagraphStreamCommit<'package>> {
        Ok(self.prepare()?.commit())
    }
}

/// Prepared semantic and physical plan for one authored stream append.
pub struct ParagraphStreamPlan<'package> {
    splice: litchi_opc::source_backed::SourcePartSplicePlan<'package>,
    source: tail_append::SourceProof,
    candidate: tail_append::CandidateProof,
    authored: AuthoredStreamProof,
    limits: ParagraphStreamLimits,
    source_artifact_fingerprint: SourceArtifactFingerprint,
    source_artifact_length: u64,
    replay: Arc<dyn AuthoredReplayHandle>,
    main_part: String,
    expected_candidate_archive: Option<patch::ArtifactProof>,
}

impl fmt::Debug for ParagraphStreamPlan<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ParagraphStreamPlan")
            .field("source", &self.source)
            .field("candidate", &self.candidate)
            .field("authored", &self.authored)
            .field("limits", &self.limits)
            .field(
                "source_artifact_fingerprint",
                &self.source_artifact_fingerprint,
            )
            .field("source_artifact_length", &self.source_artifact_length)
            .finish_non_exhaustive()
    }
}

impl<'package> ParagraphStreamPlan<'package> {
    /// Borrow the compact source semantic proof.
    #[must_use]
    pub const fn source_proof(&self) -> tail_append::SourceProof {
        self.source
    }

    /// Borrow the compact candidate semantic proof.
    #[must_use]
    pub const fn candidate_proof(&self) -> tail_append::CandidateProof {
        self.candidate
    }

    /// Borrow the authenticated authored proof.
    #[must_use]
    pub const fn authored_proof(&self) -> AuthoredStreamProof {
        self.authored
    }

    /// Return the compact source/candidate physical splice proof.
    #[must_use]
    pub const fn splice_proof(&self) -> SourcePartSpliceProof {
        self.splice.proof()
    }

    /// Return the raw source archive fingerprint captured during preparation.
    #[must_use]
    pub const fn source_artifact_fingerprint(&self) -> SourceArtifactFingerprint {
        self.source_artifact_fingerprint
    }

    /// Return the raw source archive length captured during preparation.
    #[must_use]
    pub const fn source_artifact_length(&self) -> u64 {
        self.source_artifact_length
    }

    /// Return the stream policy retained by this plan.
    #[must_use]
    pub const fn limits(&self) -> ParagraphStreamLimits {
        self.limits
    }

    /// Borrow the replay handle for the durable patch owner.
    #[must_use]
    pub(super) fn replay_handle(&self) -> &dyn AuthoredReplayHandle {
        self.replay.as_ref()
    }

    /// Return the optional durable authored provider reference.
    #[must_use]
    pub fn durable_replay_reference(&self) -> Option<AuthoredReplayReference> {
        self.replay.durable_reference()
    }

    /// Consume this plan into a named commit product.
    #[must_use]
    pub fn commit(self) -> ParagraphStreamCommit<'package> {
        ParagraphStreamCommit { plan: self }
    }

    /// Publish this plan directly to a sequential sink.
    pub fn write_to_stream(self, writer: impl Write) -> Result<ParagraphStreamPublication> {
        self.commit().write_to_stream(writer)
    }

    /// Atomically publish this plan to a filesystem path.
    ///
    /// The path publication uses the same consuming commit and source-backed
    /// splice as [`Self::write_to_stream`].  The existing splice callback
    /// authenticates the source, candidate, replay, budget, and cancellation
    /// state while writing the complete candidate to an atomic temporary
    /// sibling; only then does the filesystem layer synchronize and replace
    /// the destination.  This is not a compare-and-swap or a continuous source
    /// watch, so source freshness remains the callback's existing contract.
    /// A successful call returns the same source, candidate, replay, and
    /// inverse evidence as stream publication.
    pub fn write_to_path(self, path: impl AsRef<Path>) -> Result<ParagraphStreamPublication> {
        self.commit().write_to_path(path)
    }
}

/// Named consuming commit product for one stream append.
pub struct ParagraphStreamCommit<'package> {
    plan: ParagraphStreamPlan<'package>,
}

impl<'package> ParagraphStreamCommit<'package> {
    /// Borrow the prepared source proof.
    #[must_use]
    pub const fn source_proof(&self) -> tail_append::SourceProof {
        self.plan.source
    }

    /// Borrow the prepared authored proof.
    #[must_use]
    pub const fn authored_proof(&self) -> AuthoredStreamProof {
        self.plan.authored
    }

    /// Publish to a sequential sink and retain an exact immediate inverse.
    pub fn write_to_stream(self, writer: impl Write) -> Result<ParagraphStreamPublication> {
        publish_stream(self, writer)
    }

    /// Atomically publish this consuming commit to a filesystem path.
    ///
    /// OPC filesystem failures are returned as [`Error::Opc`], including
    /// [`litchi_opc::error::OpcError::Committed`] when replacement succeeded
    /// but the parent directory could not be synchronized.  `Committed` means
    /// the destination has already been replaced; this method returns no
    /// publication or inverse product in that case, so callers must inspect
    /// the output and must not blindly retry.  A source, candidate, replay,
    /// budget, cancellation, or inverse-proof failure from the splice callback
    /// is passed through unchanged and leaves the destination untouched.
    pub fn write_to_path(self, path: impl AsRef<Path>) -> Result<ParagraphStreamPublication> {
        publish_stream_to_path(self, path.as_ref())
    }
}

/// Successful stream publication and immediate exact-inverse authorization.
#[derive(Debug)]
pub struct ParagraphStreamPublication {
    source: tail_append::SourceProof,
    candidate: tail_append::CandidateProof,
    authored: AuthoredStreamProof,
    splice: SourcePartSplicePublication,
    source_artifact_fingerprint: SourceArtifactFingerprint,
    source_artifact_length: u64,
    limits: ParagraphStreamLimits,
    main_part: String,
    replay_reference: Option<AuthoredReplayReference>,
}

impl ParagraphStreamPublication {
    /// Borrow source semantic proof.
    #[must_use]
    pub const fn source_proof(&self) -> tail_append::SourceProof {
        self.source
    }

    /// Borrow candidate semantic proof.
    #[must_use]
    pub const fn candidate_proof(&self) -> tail_append::CandidateProof {
        self.candidate
    }

    /// Borrow authored stream proof.
    #[must_use]
    pub const fn authored_proof(&self) -> AuthoredStreamProof {
        self.authored
    }

    /// Return the published raw archive fingerprint.
    #[must_use]
    pub const fn candidate_artifact_fingerprint(&self) -> SourceArtifactFingerprint {
        self.splice.candidate_artifact_fingerprint()
    }

    /// Return the exact number of bytes in the published archive.
    #[must_use]
    pub const fn candidate_artifact_length(&self) -> u64 {
        self.splice.candidate_artifact_len()
    }

    /// Return the raw source archive fingerprint captured before publication.
    #[must_use]
    pub const fn source_artifact_fingerprint(&self) -> SourceArtifactFingerprint {
        self.source_artifact_fingerprint
    }

    /// Return the raw source archive length captured before publication.
    #[must_use]
    pub const fn source_artifact_length(&self) -> u64 {
        self.source_artifact_length
    }

    /// Build the compact forward patch authorized by this publication.
    pub fn patch(&self) -> patch::Result<patch::ParagraphStreamPatch> {
        let source_archive = patch::ArtifactProof::new(
            self.source_artifact_length,
            self.source_artifact_fingerprint.into_sha256(),
        );
        let candidate_archive = patch::ArtifactProof::new(
            self.candidate_artifact_length(),
            self.candidate_artifact_fingerprint().into_sha256(),
        );
        patch::ParagraphStreamPatch::from_tail_append_proofs(
            patch::PatchLimits::from_stream_limits(&self.limits),
            self.main_part.clone(),
            self.source,
            self.authored,
            self.candidate,
            source_archive,
            candidate_archive,
            self.replay_reference.clone(),
        )
    }

    /// Alias for callers that require an explicit durable-patch product.
    pub fn durable_patch(&self) -> patch::Result<patch::ParagraphStreamPatch> {
        self.patch()
    }

    /// Restore the exact original archive after authenticating the current candidate.
    pub fn write_inverse_to_stream(
        &self,
        current: &Package,
        writer: impl Write,
    ) -> Result<(), Error> {
        self.splice
            .write_inverse_to_stream(&current.package, writer)
            .map_err(Error::Opc)
    }
}

impl Package {
    /// Start a replayable source-backed append of plain authored paragraphs.
    #[must_use]
    pub fn tail_append_plain_paragraphs<S>(
        &self,
        source: S,
        limits: ParagraphStreamLimits,
    ) -> ParagraphStreamEdit<'_, S>
    where
        S: ReplayableParagraphSource + 'static,
    {
        ParagraphStreamEdit {
            package: self,
            input: Arc::new(DeterministicInput {
                source: Arc::new(source),
            }),
            limits,
            cancellation: None,
            expected: None,
            marker: PhantomData,
        }
    }

    /// Start a stream append whose producer runs once into an explicit replay store.
    pub fn tail_append_plain_paragraphs_from_producer<P, R>(
        &self,
        producer: P,
        replay: R,
        limits: ParagraphStreamLimits,
    ) -> Result<ParagraphStreamEdit<'_, R::Handle>, Error>
    where
        P: OneShotParagraphProducer + 'static,
        R: AuthoredReplayStore + 'static,
        R::Handle: 'static,
    {
        limits.validate()?;
        Ok(ParagraphStreamEdit {
            package: self,
            input: Arc::new(ProducerInput {
                producer_and_store: Mutex::new(Some((producer, replay))),
            }),
            limits,
            cancellation: None,
            expected: None,
            marker: PhantomData,
        })
    }

    /// Reopen and apply a durable forward patch through an explicit authored
    /// replay resolver. All semantic and source-archive proofs are checked
    /// before the supplied sink is touched; the candidate archive identity is
    /// checked against the retained patch identity as the physical publication
    /// completes.
    pub fn apply_tail_append_stream_patch<R, W>(
        &self,
        durable: &patch::ParagraphStreamPatch,
        resolver: &R,
        writer: W,
    ) -> Result<ParagraphStreamPublication>
    where
        R: patch::AuthoredReplayResolver,
        W: Write,
    {
        let limits = ParagraphStreamLimits::from_patch_limits(durable.limits())?;
        let expected = PatchExpectations::from_patch(durable);
        let expected_reference = expected
            .replay_reference
            .clone()
            .ok_or(Error::Patch(patch::PatchError::MissingReplayProvider))?;
        let source_archive = expected.source.archive;
        check_package_execution_context(self)?;
        ensure_current_archive(self, source_archive)?;
        check_package_execution_context(self)?;
        let replay = match durable.resolve_replay(resolver) {
            Ok(replay) => replay,
            Err(error) => {
                // A callback failure must not mask a simultaneous change to
                // the source that authorizes this patch.
                ensure_current_archive(self, source_archive)?;
                return Err(Error::Patch(error));
            },
        };
        check_package_execution_context(self)?;
        let replay_input: Arc<dyn AuthoredInput> = Arc::new(ResolvedInput {
            handle: replay,
            expected: expected.authored,
            expected_reference,
        });
        let make_edit = || -> ParagraphStreamEdit<'_, ()> {
            ParagraphStreamEdit {
                package: self,
                input: Arc::clone(&replay_input),
                limits,
                cancellation: None,
                expected: Some(expected.clone()),
                marker: PhantomData,
            }
        };
        let plan = make_edit().prepare()?;
        ensure_current_archive(self, source_archive)?;
        plan.write_to_stream(writer)
    }

    /// Apply a durable stream patch using the shorter transaction-family
    /// name.
    pub fn apply_tail_append_patch<R, W>(
        &self,
        durable: &patch::ParagraphStreamPatch,
        resolver: &R,
        writer: W,
    ) -> Result<ParagraphStreamPublication>
    where
        R: patch::AuthoredReplayResolver,
        W: Write,
    {
        self.apply_tail_append_stream_patch(durable, resolver, writer)
    }
}

fn ensure_current_archive(package: &Package, expected: patch::ArtifactProof) -> Result<(), Error> {
    let artifact = package.package.source_artifact();
    let actual = patch::ArtifactProof::new(
        artifact.len(),
        artifact.fingerprint().map_err(Error::Opc)?.into_sha256(),
    );
    if actual != expected {
        return Err(Error::Opc(
            litchi_opc::error::OpcError::SourceArtifactMismatch {
                artifact: "source",
                field: "length or SHA-256",
            },
        ));
    }
    Ok(())
}

fn check_package_execution_context(package: &Package) -> Result<(), Error> {
    if let Some(context) = package.package.execution_context() {
        context
            .check()
            .map_err(|error| Error::TailAppend(tail_append::Error::Execution(error)))?;
    }
    Ok(())
}

#[derive(Debug)]
enum CandidateError {
    Tail(tail_append::Error),
    Replay(AuthoredReplayError),
}

fn map_candidate_error(error: CandidateError) -> Error {
    match error {
        CandidateError::Tail(error) => Error::TailAppend(error),
        CandidateError::Replay(error) => Error::Replay(error),
    }
}

fn prepare_stream<'package, S>(
    edit: ParagraphStreamEdit<'package, S>,
) -> Result<ParagraphStreamPlan<'package>> {
    edit.limits.validate()?;
    let ParagraphStreamEdit {
        package,
        input,
        limits,
        cancellation,
        expected,
        marker: _,
    } = edit;
    let context = package.package.execution_context();
    if let Some(context) = context.as_ref() {
        context
            .check()
            .map_err(|error| Error::TailAppend(tail_append::Error::Execution(error)))?;
    }
    if let Some(token) = cancellation.as_ref() {
        token
            .check()
            .map_err(|error| Error::TailAppend(tail_append::Error::Execution(error)))?;
    }
    tail_append::validate_topology(
        package,
        limits.source,
        context.as_ref(),
        cancellation.as_ref(),
    )?;
    let source_version = package
        .package
        .source_version()
        .map_err(|error| Error::TailAppend(tail_append::Error::Document(error.into())))?;
    let main = package
        .package
        .main_document_part()
        .map_err(|error| Error::TailAppend(tail_append::Error::Document(error.into())))?;
    if let Some(expected) = expected.as_ref() {
        if main.partname().as_str() != expected.source.main_part.as_str() {
            return Err(Error::Patch(patch::PatchError::InvalidFacts(
                "durable patch main part does not match the package",
            )));
        }
    }
    crate::package::validate_document_main_content_type(main.content_type())
        .map_err(tail_append::Error::Document)
        .map_err(Error::TailAppend)?;
    let source_len = main
        .declared_uncompressed_size()
        .map_err(|error| Error::TailAppend(tail_append::Error::Document(error.into())))?;
    check_limit(
        "source XML bytes",
        source_len,
        limits.source.max_source_xml_bytes,
    )?;
    let source_scan = tail_append::scan_main_part(
        &main,
        source_len,
        limits.source,
        context.as_ref(),
        cancellation.as_ref(),
        None,
    )?;
    let source_artifact = package.package.source_artifact();
    let source_artifact_length = source_artifact.len();
    let source_artifact_fingerprint = source_artifact.fingerprint().map_err(Error::Opc)?;
    if let Some(expected) = expected.as_ref() {
        let actual = patch::ArtifactProof::new(
            source_artifact_length,
            source_artifact_fingerprint.into_sha256(),
        );
        if actual != expected.source.archive {
            return Err(Error::Patch(patch::PatchError::InvalidFacts(
                "durable patch source archive does not match the package",
            )));
        }
    }
    let authored = input.seal(
        source_scan.strict_namespace,
        limits,
        context.as_ref(),
        cancellation.as_ref(),
    )?;
    let authored_proof = authored.proof();
    check_limit(
        "authored XML bytes",
        authored_proof.encoded_xml_bytes,
        limits.max_authored_xml_bytes,
    )?;
    check_limit(
        "fragment bytes",
        authored_proof.encoded_xml_bytes,
        limits.source.max_fragment_bytes,
    )?;
    let candidate_len = source_len
        .checked_add(authored_proof.encoded_xml_bytes)
        .ok_or(Error::Limit {
            resource: "candidate XML bytes",
            actual: u64::MAX,
            maximum: limits.source.max_candidate_xml_bytes,
        })?;
    check_limit(
        "candidate XML bytes",
        candidate_len,
        limits.source.max_candidate_xml_bytes,
    )?;

    let candidate_scan = main.with_verified_decoded_reader(|reader| {
        let replay_reader = authored.open().map_err(CandidateError::Replay)?;
        let mut splice = ReplaySpliceReader::new(
            reader,
            source_len,
            source_scan.insertion_offset,
            replay_reader,
            limits.max_replay_window_bytes,
            context.as_ref(),
        )
        .map_err(CandidateError::Replay)?;
        let facts = tail_append::scan_reader_checked(
            &mut splice,
            candidate_len,
            limits.source,
            context.as_ref(),
            cancellation.as_ref(),
            Some((
                source_scan.insertion_offset,
                authored_proof.encoded_xml_bytes,
                authored_proof.paragraph_count,
            )),
        )
        .map_err(CandidateError::Tail)?;
        let pass = splice.finish_replay().map_err(CandidateError::Replay)?;
        if pass.proof() != authored_proof {
            return Err(CandidateError::Replay(AuthoredReplayError::Changed));
        }
        Ok(facts)
    });
    let candidate_scan = match candidate_scan {
        Ok(facts) => facts,
        Err(VerifiedDecodedReaderError::Callback(error)) => {
            return Err(map_candidate_error(error));
        },
        Err(VerifiedDecodedReaderError::Opc { error, .. }) => {
            return Err(Error::TailAppend(tail_append::Error::Document(
                error.into(),
            )));
        },
        Err(_) => {
            return Err(Error::Invalid(
                "verified candidate reader returned an unknown failure",
            ));
        },
    };
    validate_stream_candidate(&source_scan, &candidate_scan, authored_proof, limits.source)?;
    let source_proof = source_scan.source_proof(source_version);
    let candidate_proof = candidate_scan.candidate_proof();
    if let Some(expected) = expected.as_ref() {
        if authored_proof != expected.authored
            || !source_proof_matches(&expected.source, source_proof)
            || !candidate_proof_matches(&expected.candidate, candidate_proof, authored_proof)
        {
            return Err(Error::Patch(patch::PatchError::InvalidFacts(
                "durable patch semantic proof does not match the reopened package",
            )));
        }
    }
    let splice_proof = SourcePartSpliceProof {
        source_version,
        source_len,
        source_sha256: source_scan.sha256,
        insertion_offset: source_scan.insertion_offset,
        fragment_len: authored_proof.encoded_xml_bytes,
        fragment_sha256: authored_proof.encoded_sha256,
        candidate_len,
        candidate_sha256: candidate_scan.sha256,
    };
    let splice_limits = make_stream_splice_limits(limits)?;
    let opc_replay: Arc<dyn SourcePartSpliceReplay> =
        Arc::new(OpcReplayAdapter::new(Arc::clone(&authored)));
    let splice = package
        .package
        .prepare_source_part_splice_with_replay(
            main.partname(),
            splice_proof,
            opc_replay,
            splice_limits,
        )
        .map_err(Error::Opc)?;
    package
        .package
        .source_version()
        .map_err(|error| Error::TailAppend(tail_append::Error::Document(error.into())))?;
    if let Some(context) = context.as_ref() {
        context
            .check()
            .map_err(|error| Error::TailAppend(tail_append::Error::Execution(error)))?;
    }
    if let Some(token) = cancellation.as_ref() {
        token
            .check()
            .map_err(|error| Error::TailAppend(tail_append::Error::Execution(error)))?;
    }
    Ok(ParagraphStreamPlan {
        splice,
        source: source_proof,
        candidate: candidate_proof,
        authored: authored_proof,
        limits,
        source_artifact_fingerprint,
        source_artifact_length,
        replay: authored,
        main_part: main.partname().as_str().to_owned(),
        expected_candidate_archive: expected
            .as_ref()
            .map(|expectation| expectation.candidate.archive),
    })
}

fn make_stream_splice_limits(limits: ParagraphStreamLimits) -> Result<SourcePartSpliceLimits> {
    let splice = tail_append::make_splice_limits_for_stream(limits.source)?;
    Ok(splice.with_max_authored_replay_window_bytes(limits.max_replay_window_bytes))
}

fn validate_stream_candidate(
    source: &tail_append::ScanFacts,
    candidate: &tail_append::ScanFacts,
    authored: AuthoredStreamProof,
    limits: tail_append::Limits,
) -> Result<(), Error> {
    let expected_len = source
        .len
        .checked_add(authored.encoded_xml_bytes)
        .ok_or(Error::Limit {
            resource: "candidate XML bytes",
            actual: u64::MAX,
            maximum: limits.max_candidate_xml_bytes,
        })?;
    let expected_anchor = source
        .insertion_offset
        .checked_add(authored.encoded_xml_bytes)
        .ok_or(Error::Limit {
            resource: "candidate XML bytes",
            actual: u64::MAX,
            maximum: limits.max_candidate_xml_bytes,
        })?;
    let expected_paragraphs = source
        .paragraph_count
        .checked_add(authored.paragraph_count)
        .ok_or(Error::Limit {
            resource: "paragraphs",
            actual: u64::MAX,
            maximum: limits.max_paragraphs,
        })?;
    if candidate.len != expected_len
        || candidate.paragraph_count != expected_paragraphs
        || candidate.generated_count != authored.paragraph_count
        || !candidate.generated_once
        || candidate.generated_offset != source.insertion_offset
        || candidate.insertion_offset != expected_anchor
        || candidate.strict_namespace != source.strict_namespace
        || candidate.sect_pr_len != source.sect_pr_len
        || candidate.sect_pr_sha256 != source.sect_pr_sha256
        || candidate.event_count <= source.event_count
        || candidate.max_depth < source.max_depth
    {
        return Err(Error::Invalid(
            "candidate semantic proof does not match the authored stream insertion",
        ));
    }
    Ok(())
}

fn source_proof_matches(
    expected: &patch::StreamSourceProof,
    actual: tail_append::SourceProof,
) -> bool {
    expected.source_len == actual.source_len
        && expected.source_sha256 == actual.source_sha256
        && expected.insertion_offset == actual.insertion_offset
        && expected.paragraph_count == actual.paragraph_count
        && expected.event_count == actual.event_count
        && expected.max_depth == actual.max_depth
        && expected.strict_namespace == actual.strict_namespace
        && expected.sect_pr_len == actual.sect_pr_len
        && expected.sect_pr_sha256 == actual.sect_pr_sha256
}

fn candidate_proof_matches(
    expected: &patch::StreamCandidateProof,
    actual: tail_append::CandidateProof,
    authored: AuthoredStreamProof,
) -> bool {
    expected.candidate_len == actual.candidate_len
        && expected.candidate_sha256 == actual.candidate_sha256
        && expected.paragraph_count == actual.paragraph_count
        && expected.event_count == actual.event_count
        && expected.max_depth == actual.max_depth
        && expected.generated_offset == actual.generated_offset
        && expected.generated_paragraph_count == authored.paragraph_count
        && expected.generated_once == actual.generated_once
        && expected.sect_pr_len == actual.sect_pr_len
        && expected.sect_pr_sha256 == actual.sect_pr_sha256
}

fn publish_stream(
    commit: ParagraphStreamCommit<'_>,
    writer: impl Write,
) -> Result<ParagraphStreamPublication> {
    let ParagraphStreamCommit { plan } = commit;
    let replay_reference = plan.replay_handle().durable_reference();
    let ParagraphStreamPlan {
        splice,
        source,
        candidate,
        authored,
        limits,
        source_artifact_fingerprint,
        source_artifact_length,
        replay: _,
        main_part,
        expected_candidate_archive,
    } = plan;
    let splice = match expected_candidate_archive {
        Some(expected) => splice
            .write_to_stream_with_expected_artifact(
                writer,
                expected.len,
                SourceArtifactFingerprint::from_sha256(expected.sha256),
            )
            .map_err(Error::Opc)?,
        None => splice.write_to_stream(writer).map_err(Error::Opc)?,
    };
    Ok(ParagraphStreamPublication {
        source,
        candidate,
        authored,
        splice,
        source_artifact_fingerprint,
        source_artifact_length,
        limits,
        main_part,
        replay_reference,
    })
}

fn publish_stream_to_path(
    commit: ParagraphStreamCommit<'_>,
    path: &Path,
) -> Result<ParagraphStreamPublication> {
    let mut publication = None;
    litchi_opc::atomic::replace_with::<Error>(path, |temporary| {
        let published = publish_stream(commit, temporary)?;
        publication = Some(published);
        Ok(())
    })?;
    publication.ok_or(Error::Invalid(
        "atomic stream publication completed without a publication product",
    ))
}

struct ReplaySpliceReader<'a> {
    source: &'a mut dyn BufRead,
    source_len: u64,
    insertion: u64,
    source_position: u64,
    replay: Option<Box<dyn AuthoredReplayReader + 'a>>,
    replay_buffer: Vec<u8>,
    replay_valid_len: usize,
    replay_position: usize,
    replay_eof: bool,
    _replay_window_reservation: Option<Reservation>,
}

impl<'a> ReplaySpliceReader<'a> {
    fn new(
        source: &'a mut dyn BufRead,
        source_len: u64,
        insertion: u64,
        replay: Box<dyn AuthoredReplayReader + 'a>,
        window: u64,
        context: Option<&ExecutionContext>,
    ) -> Result<Self, AuthoredReplayError> {
        if insertion > source_len {
            return Err(AuthoredReplayError::Invalid(
                "replay insertion offset is outside the source",
            ));
        }
        let window = usize::try_from(window).map_err(|_| AuthoredReplayError::Limit {
            resource: "replay window bytes",
            actual: window,
            maximum: usize::MAX as u64,
        })?;
        if window == 0 {
            return Err(AuthoredReplayError::Limit {
                resource: "replay window bytes",
                actual: 0,
                maximum: 0,
            });
        }
        let replay_window_reservation = context
            .map(|context| {
                context
                    .reserve(Resource::Memory, window as u64)
                    .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))
            })
            .transpose()?;
        let mut replay_buffer = Vec::new();
        replay_buffer
            .try_reserve_exact(window)
            .map_err(|_| AuthoredReplayError::Store("replay window allocation failed"))?;
        if replay_buffer.capacity() != window {
            drop(replay_window_reservation);
            return Err(AuthoredReplayError::Store(
                "replay window allocation exceeded its reservation",
            ));
        }
        replay_buffer.resize(window, 0);
        Ok(Self {
            source,
            source_len,
            insertion,
            source_position: 0,
            replay: Some(replay),
            replay_buffer,
            replay_valid_len: 0,
            replay_position: 0,
            replay_eof: false,
            _replay_window_reservation: replay_window_reservation,
        })
    }

    fn finish_replay(&mut self) -> Result<AuthoredPassProof, AuthoredReplayError> {
        if !self.replay_eof {
            return Err(AuthoredReplayError::Invalid(
                "candidate scanner did not consume the complete authored replay",
            ));
        }
        self.replay
            .take()
            .ok_or(AuthoredReplayError::Invalid(
                "authored replay was already finished",
            ))?
            .finish()
    }
}

impl Read for ReplaySpliceReader<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let available = self.fill_buf()?;
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

impl BufRead for ReplaySpliceReader<'_> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.source_position < self.insertion {
            let remaining =
                usize::try_from(self.insertion - self.source_position).unwrap_or(usize::MAX);
            let available = self.source.fill_buf()?;
            return Ok(&available[..available.len().min(remaining)]);
        }
        if !self.replay_eof {
            if self.replay_position < self.replay_valid_len {
                return Ok(&self.replay_buffer[self.replay_position..self.replay_valid_len]);
            }
            self.replay_position = 0;
            self.replay_valid_len = 0;
            let mut replay = self
                .replay
                .take()
                .ok_or_else(|| io::Error::other("authored replay reader was consumed"))?;
            let count = replay.read(&mut self.replay_buffer);
            match count {
                Ok(0) => {
                    self.replay = Some(replay);
                    self.replay_eof = true;
                },
                Ok(count) => {
                    if count > self.replay_buffer.len() {
                        self.replay = Some(replay);
                        return Err(map_replay_io(AuthoredReplayError::Invalid(
                            "authored replay reader returned more bytes than its window",
                        )));
                    }
                    self.replay = Some(replay);
                    self.replay_valid_len = count;
                    return Ok(&self.replay_buffer[..count]);
                },
                Err(error) => {
                    self.replay = Some(replay);
                    return Err(error);
                },
            }
        }
        if self.source_position >= self.source_len {
            return Ok(&[]);
        }
        let remaining =
            usize::try_from(self.source_len - self.source_position).unwrap_or(usize::MAX);
        let available = self.source.fill_buf()?;
        Ok(&available[..available.len().min(remaining)])
    }

    fn consume(&mut self, amount: usize) {
        if self.source_position < self.insertion {
            let remaining =
                usize::try_from(self.insertion - self.source_position).unwrap_or(usize::MAX);
            let count = amount.min(remaining);
            if count != 0 {
                self.source.consume(count);
                self.source_position = self.source_position.saturating_add(count as u64);
            }
            return;
        }
        if !self.replay_eof && self.replay_position < self.replay_valid_len {
            let available = self.replay_valid_len - self.replay_position;
            self.replay_position += amount.min(available);
            return;
        }
        let remaining = usize::try_from(self.source_len.saturating_sub(self.source_position))
            .unwrap_or(usize::MAX);
        let count = amount.min(remaining);
        if count != 0 {
            self.source.consume(count);
            self.source_position = self.source_position.saturating_add(count as u64);
        }
    }
}

#[cfg(test)]
mod authenticated_reader_tests {
    use super::*;

    #[derive(Clone, Copy)]
    enum ReadFailure {
        Overreported,
        ExceedsProof,
    }

    struct FakeReader {
        failure: ReadFailure,
        proof: AuthoredPassProof,
    }

    impl Read for FakeReader {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            match self.failure {
                ReadFailure::Overreported => Ok(output.len().saturating_add(1)),
                ReadFailure::ExceedsProof => {
                    if let Some(first) = output.first_mut() {
                        *first = b'x';
                        Ok(1)
                    } else {
                        Ok(0)
                    }
                },
            }
        }
    }

    impl AuthoredReplayReader for FakeReader {
        fn finish(self: Box<Self>) -> Result<AuthoredPassProof, AuthoredReplayError> {
            Ok(self.proof)
        }
    }

    fn proof(encoded_xml_bytes: u64) -> AuthoredStreamProof {
        AuthoredStreamProof {
            strict_namespace: false,
            paragraph_count: 1,
            event_count: 3,
            text_bytes: 0,
            encoded_xml_bytes,
            event_sha256: [0; 32],
            encoded_sha256: [0; 32],
        }
    }

    fn reader(failure: ReadFailure, expected: AuthoredStreamProof) -> AuthenticatedReader<'static> {
        let inner = FakeReader {
            failure,
            proof: AuthoredPassProof(expected),
        };
        AuthenticatedReader {
            inner: Some(Box::new(inner)),
            identity: None,
            durable_reference: None,
            expected,
            bytes: 0,
            hash: Sha256::new(),
            finished: false,
            poisoned: false,
        }
    }

    fn authored_error(error: &io::Error) -> Option<&AuthoredReplayError> {
        error
            .get_ref()
            .and_then(|source| source.downcast_ref::<AuthoredReplayError>())
    }

    #[test]
    fn authenticated_reader_rejects_overreported_count_and_poison_repeats() {
        let expected = proof(4);
        let mut reader = reader(ReadFailure::Overreported, expected);
        let first = reader.read(&mut [0; 4]).expect_err("overreported read");
        assert!(matches!(
            authored_error(&first),
            Some(AuthoredReplayError::Invalid(_))
        ));
        assert!(reader.read(&mut [0; 4]).is_err());
    }

    #[test]
    fn authenticated_reader_rejects_bytes_beyond_proof_before_hashing() {
        let expected = proof(0);
        let mut reader = reader(ReadFailure::ExceedsProof, expected);
        let first = reader.read(&mut [0; 1]).expect_err("excess proof bytes");
        assert!(matches!(
            authored_error(&first),
            Some(AuthoredReplayError::Changed)
        ));
        assert!(reader.read(&mut [0; 1]).is_err());
    }
}
