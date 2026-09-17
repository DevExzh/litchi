//! Part payload storage, including payloads that are still in the source ZIP.
//!
//! An [`OpcPackage`](crate::package::OpcPackage) opened from an **owned**
//! source archive keeps that archive for exact and targeted publication. Its
//! parts therefore do not need a decompressed copy of every payload before a
//! caller has asked for one: the payload can stay in the retained archive and
//! be inflated on first access.
//!
//! [`PartPayload`] is the storage that makes that possible. A payload is
//! either [`PartPayload::Ready`] — an already materialized allocation, which
//! is what borrowed ingress, streamed ingress and in-memory authoring produce
//! — or [`PartPayload::Deferred`], a handle to one member of the retained
//! source archive that inflates on first access and at most once.
//!
//! See [ADR 0030](../../../../docs/adr/0030-lazy-opc-part-decode.md) for the
//! contract this implements, in particular which refusals stay at `open()` and
//! which become first-access refusals.

use std::sync::Arc;
use std::sync::OnceLock;
use std::sync::atomic::{AtomicU64, Ordering};

use soapberry_zip::office::IndexedArchive;

use crate::error::{OpcError, Result, replicate_deferred_error};
use crate::limits::{ReadLimits, ReadResource};

/// The retained owned source archive that deferred payloads inflate from.
///
/// One value is shared by every deferred part of a single package open, so the
/// ZIP index is built at most once and the aggregate actual-byte charge is a
/// single running total, exactly as the eager reader's single pass computes
/// it.
#[derive(Debug)]
pub(crate) struct DeferredPartSource {
    /// The retained source archive. This is the same allocation the package
    /// holds as `source_archive`, so a deferred package retains one copy of
    /// the compressed bytes, not two.
    bytes: Arc<Vec<u8>>,
    limits: ReadLimits,
    /// The ZIP index over `bytes`, built on the first decode and never again.
    /// A package that decodes nothing never builds it.
    index: OnceLock<std::result::Result<IndexedArchive<Arc<Vec<u8>>>, OpcError>>,
    /// Running total of **actual** inflated part bytes, charged against
    /// `max_total_part_bytes`. The eager reader charges the same quantity in
    /// its bulk loop; this charges it as parts are decoded.
    inflated_bytes: AtomicU64,
    /// Parts inflated through this source. Retained for measurement and for
    /// the regression tests that assert how many parts an operation reads.
    inflated_parts: AtomicU64,
}

impl DeferredPartSource {
    pub(crate) fn new(bytes: Arc<Vec<u8>>, limits: ReadLimits) -> Self {
        Self {
            bytes,
            limits,
            index: OnceLock::new(),
            inflated_bytes: AtomicU64::new(0),
            inflated_parts: AtomicU64::new(0),
        }
    }

    /// Parts inflated and bytes inflated through this source so far.
    pub(crate) fn counters(&self) -> (u64, u64) {
        (
            self.inflated_parts.load(Ordering::Relaxed),
            self.inflated_bytes.load(Ordering::Relaxed),
        )
    }

    /// The ZIP index over the retained archive, built at most once.
    fn index(&self) -> Result<&IndexedArchive<Arc<Vec<u8>>>> {
        let end_offset = self.bytes.len() as u64;
        let limits = self.limits;
        let bytes = &self.bytes;
        self.index
            .get_or_init(|| {
                IndexedArchive::from_reader_with_limits(
                    Arc::clone(bytes),
                    end_offset,
                    limits.zip_limits(),
                )
                .map_err(OpcError::from)
            })
            .as_ref()
            .map_err(replicate_deferred_error)
    }

    /// Charge one decoded payload against the aggregate actual-byte limit.
    ///
    /// This mirrors `pkgreader`'s `checked_add` for
    /// [`ReadResource::TotalPartBytes`], including the overflow observation
    /// value, so a lazy package's refusal is the value the eager path
    /// produces.
    fn charge_inflated(&self, bytes: u64) -> Result<()> {
        let maximum = self.limits.max_total_part_bytes();
        let mut current = self.inflated_bytes.load(Ordering::Relaxed);
        loop {
            let actual = current.checked_add(bytes).ok_or(OpcError::ReadLimit {
                resource: ReadResource::TotalPartBytes,
                actual: u64::MAX,
                maximum,
            })?;
            if actual > maximum {
                return Err(OpcError::ReadLimit {
                    resource: ReadResource::TotalPartBytes,
                    actual,
                    maximum,
                });
            }
            match self.inflated_bytes.compare_exchange_weak(
                current,
                actual,
                Ordering::Relaxed,
                Ordering::Relaxed,
            ) {
                Ok(_) => {
                    self.inflated_parts.fetch_add(1, Ordering::Relaxed);
                    return Ok(());
                },
                Err(observed) => current = observed,
            }
        }
    }
}

/// One part's payload, still held by the retained source archive.
#[derive(Debug)]
pub(crate) struct DeferredPayload {
    source: Arc<DeferredPartSource>,
    /// The ZIP member name this payload inflates from.
    member: Box<str>,
    /// The decode outcome, recorded once. A failure is recorded too, so the
    /// refusal is stable and idempotent across repeated accesses: a limit
    /// failure that alternated with success would make the error identity
    /// depend on call order.
    cell: OnceLock<std::result::Result<Arc<Vec<u8>>, OpcError>>,
}

impl DeferredPayload {
    fn decode(&self) -> std::result::Result<Arc<Vec<u8>>, OpcError> {
        let archive = self.source.index()?;
        let blob = archive.read(&self.member).map_err(OpcError::from)?;
        let limits = self.source.limits;
        let inflated = blob.len() as u64;
        limits.check(ReadResource::PartBytes, inflated, limits.max_part_bytes())?;
        self.source.charge_inflated(inflated)?;
        Ok(Arc::new(blob))
    }

    fn force(&self) -> Result<&Arc<Vec<u8>>> {
        self.cell
            .get_or_init(|| self.decode())
            .as_ref()
            .map_err(replicate_deferred_error)
    }
}

/// A part's payload storage.
///
/// `Ready` is today's representation and is what every ingress except an
/// owned-source open produces. `Deferred` is a payload that is still in the
/// retained source archive.
#[derive(Debug, Clone)]
pub(crate) enum PartPayload {
    Ready(Arc<Vec<u8>>),
    /// Cloning shares the cell rather than copying it, so a package and its
    /// clones decode a part at most once between them and charge the
    /// aggregate budget once.
    Deferred(Arc<DeferredPayload>),
}

/// The empty payload a failed decode presents to the infallible accessors.
///
/// `Part::blob` cannot report an error. A payload whose decode failed has no
/// bytes, and every route that could publish or interpret those bytes goes
/// through a fallible accessor that returns the recorded refusal first.
fn empty_payload() -> &'static Arc<Vec<u8>> {
    static EMPTY: OnceLock<Arc<Vec<u8>>> = OnceLock::new();
    EMPTY.get_or_init(|| Arc::new(Vec::new()))
}

impl PartPayload {
    pub(crate) fn ready(bytes: Arc<Vec<u8>>) -> Self {
        Self::Ready(bytes)
    }

    pub(crate) fn deferred(source: &Arc<DeferredPartSource>, member: Box<str>) -> Self {
        Self::Deferred(Arc::new(DeferredPayload {
            source: Arc::clone(source),
            member,
            cell: OnceLock::new(),
        }))
    }

    /// Decode the payload if it is still deferred, reporting the typed
    /// refusal a failed decode recorded.
    ///
    /// # Errors
    ///
    /// Returns the refusal the first decode produced: a
    /// [`OpcError::ReadLimit`] for `PartBytes` or `TotalPartBytes`, an
    /// allocation failure, a cancellation, an I/O error, or a ZIP error for a
    /// corrupt Deflate stream or a CRC mismatch.
    pub(crate) fn force(&self) -> Result<&Arc<Vec<u8>>> {
        match self {
            Self::Ready(bytes) => Ok(bytes),
            Self::Deferred(deferred) => deferred.force(),
        }
    }

    /// The payload bytes when they are already available.
    ///
    /// This infallible observation never performs I/O or decompression. A
    /// deferred payload therefore reads as empty until the owning package has
    /// forced it through a fallible accessor. The package invariant is that a
    /// `&dyn Part` is handed to code outside this crate only after that force;
    /// keeping this method observation-only also means a recorded decode
    /// failure can never be swallowed by a hidden retry.
    pub(crate) fn bytes(&self) -> &[u8] {
        self.decoded().map_or(&[], |bytes| bytes.as_slice())
    }

    /// The payload as a shared allocation when it is already available.
    ///
    /// Like [`Self::bytes`], this method is observation-only. Internal callers
    /// that need a payload use [`Self::force`] first or arrive through a
    /// package accessor that performed the same fallible check.
    pub(crate) fn arc(&self) -> Arc<Vec<u8>> {
        self.decoded()
            .map_or_else(|| Arc::clone(empty_payload()), Arc::clone)
    }

    /// The payload if it is already available, without decoding anything.
    ///
    /// `None` means the part still holds the payload its source member
    /// carries — either because nothing has forced it or because forcing it
    /// failed. In both cases no caller has ever held the bytes, so the part
    /// cannot have been replaced and the source member is still exact.
    pub(crate) fn decoded(&self) -> Option<&Arc<Vec<u8>>> {
        match self {
            Self::Ready(bytes) => Some(bytes),
            Self::Deferred(deferred) => deferred
                .cell
                .get()
                .and_then(|outcome| outcome.as_ref().ok()),
        }
    }

    /// The source this deferred payload decodes from, for counter reporting.
    pub(crate) fn deferred_source(&self) -> Option<&Arc<DeferredPartSource>> {
        match self {
            Self::Ready(_) => None,
            Self::Deferred(deferred) => Some(&deferred.source),
        }
    }
}

/// An opaque capture of a part's payload storage.
///
/// The package captures this at open so publication can later prove a part
/// still holds the payload its source member carries. It exposes nothing on
/// its own: it is neither an archive handle nor a lock, and a caller can do
/// nothing with it but hand it back.
#[doc(hidden)]
#[derive(Debug, Clone)]
pub struct PayloadHandle(pub(crate) PartPayload);

impl PayloadHandle {
    pub(crate) fn payload(&self) -> &PartPayload {
        &self.0
    }
}
