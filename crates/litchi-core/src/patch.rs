//! Versioned, deterministic-JSON patch envelopes shared by document formats.
//!
//! This module deliberately owns transport and retention mechanics rather than
//! format semantics. Each format keeps responsibility for selector resolution,
//! read/write-set calculation, source checks, and applying an operation to its
//! own snapshot. The envelope gives those owners one durable representation
//! that can be exchanged across processes without exposing private file IDs.

use std::{
    collections::{BTreeMap, BTreeSet, VecDeque},
    fmt,
    marker::PhantomData,
    sync::Arc,
};

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use sha2::{Digest as _, Sha256};
use thiserror::Error;

/// The current durable patch wire schema.
pub const PATCH_WIRE_VERSION: u16 = 1;

const MAX_FORMAT_NAME_BYTES: usize = 128;
const MAX_OPERATION_NAME_BYTES: usize = 128;
const MAX_TARGET_BYTES: usize = 4096;

/// Typestate marker for a patch which still contains inverse operations.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Reversible;

/// Typestate marker for a redacted, forward-only patch.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ForwardOnly;

/// A content-free SHA-256 fingerprint intended only for diagnostics.
///
/// A fingerprint does not authorize patch application and does not replace a
/// format owner's exact immutable source-lineage check.
#[derive(Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct DiagnosticFingerprint([u8; 32]);

impl DiagnosticFingerprint {
    /// Fingerprints an exact byte sequence.
    #[must_use]
    pub fn of(bytes: &[u8]) -> Self {
        Self(Sha256::digest(bytes).into())
    }

    /// Returns the lowercase hexadecimal SHA-256 fingerprint.
    #[must_use]
    pub fn as_hex(&self) -> String {
        let mut text = String::with_capacity(64);
        for byte in self.0 {
            use std::fmt::Write as _;
            let _result = write!(text, "{byte:02x}");
        }
        text
    }
}

impl fmt::Debug for DiagnosticFingerprint {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_tuple("DiagnosticFingerprint")
            .field(&self.as_hex())
            .finish()
    }
}

impl fmt::Display for DiagnosticFingerprint {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.as_hex())
    }
}

/// A SHA-256 content address for one attached patch blob.
#[derive(Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct BlobId([u8; 32]);

impl BlobId {
    /// Computes the content address for `bytes`.
    #[must_use]
    pub fn of(bytes: &[u8]) -> Self {
        Self(Sha256::digest(bytes).into())
    }

    /// Returns the lowercase hexadecimal SHA-256 address.
    #[must_use]
    pub fn as_hex(&self) -> String {
        let mut text = String::with_capacity(64);
        for byte in self.0 {
            use std::fmt::Write as _;
            let _result = write!(text, "{byte:02x}");
        }
        text
    }

    fn parse_hex(text: &str) -> Result<Self, PatchError> {
        if text.len() != 64 || !text.is_ascii() {
            return Err(PatchError::InvalidBlobId);
        }
        let mut bytes = [0_u8; 32];
        for (index, pair) in text.as_bytes().chunks_exact(2).enumerate() {
            let high = hex_nibble(pair[0]).ok_or(PatchError::InvalidBlobId)?;
            let low = hex_nibble(pair[1]).ok_or(PatchError::InvalidBlobId)?;
            bytes[index] = (high << 4) | low;
        }
        if text.bytes().any(|byte| byte.is_ascii_uppercase()) {
            return Err(PatchError::InvalidBlobId);
        }
        Ok(Self(bytes))
    }
}

impl fmt::Debug for BlobId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_tuple("BlobId")
            .field(&self.as_hex())
            .finish()
    }
}

impl fmt::Display for BlobId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.as_hex())
    }
}

/// Limits applied before patch blobs are retained or decoded.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BlobLimits {
    blobs: usize,
    blob_bytes: usize,
    total_bytes: usize,
}

impl BlobLimits {
    /// Creates explicit finite limits for one patch's blob bundle.
    #[must_use]
    pub const fn new(max_blobs: usize, max_blob_bytes: usize, max_total_bytes: usize) -> Self {
        Self {
            blobs: max_blobs,
            blob_bytes: max_blob_bytes,
            total_bytes: max_total_bytes,
        }
    }

    /// Maximum number of distinct blob payloads.
    #[must_use]
    pub const fn max_blobs(self) -> usize {
        self.blobs
    }

    /// Maximum byte length of one decoded blob payload.
    #[must_use]
    pub const fn max_blob_bytes(self) -> usize {
        self.blob_bytes
    }

    /// Maximum aggregate byte length of all decoded blob payloads.
    #[must_use]
    pub const fn max_total_bytes(self) -> usize {
        self.total_bytes
    }
}

/// Finite limits applied before a durable patch is parsed or retained.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "the type is re-exported as the public bounds for durable patches"
)]
pub struct PatchLimits {
    blobs: BlobLimits,
    json_bytes: usize,
    operations: usize,
    value_depth: usize,
    value_string_bytes: usize,
    payload_bytes: usize,
}

impl PatchLimits {
    /// Creates explicit finite limits for one untrusted durable patch.
    #[must_use]
    pub const fn new(
        blobs: BlobLimits,
        max_json_bytes: usize,
        max_operations: usize,
        max_value_depth: usize,
        max_value_string_bytes: usize,
        max_payload_bytes: usize,
    ) -> Self {
        Self {
            blobs,
            json_bytes: max_json_bytes,
            operations: max_operations,
            value_depth: max_value_depth,
            value_string_bytes: max_value_string_bytes,
            payload_bytes: max_payload_bytes,
        }
    }

    /// Bounds for content-addressed blob payloads.
    #[must_use]
    pub const fn blobs(self) -> BlobLimits {
        self.blobs
    }
}

/// Bounded, content-addressed payloads referenced by semantic operations.
#[derive(Clone)]
pub struct BlobBundle {
    limits: BlobLimits,
    bytes: usize,
    entries: BTreeMap<BlobId, Arc<[u8]>>,
}

impl BlobBundle {
    /// Creates an empty bundle constrained by `limits`.
    #[must_use]
    pub fn new(limits: BlobLimits) -> Self {
        Self {
            limits,
            bytes: 0,
            entries: BTreeMap::new(),
        }
    }

    /// Adds bytes once and returns their content address.
    ///
    /// Re-inserting an identical blob is a no-op and does not consume another
    /// count or byte allowance.
    ///
    /// # Errors
    ///
    /// Returns [`PatchError::BlobLimit`] when a finite bound would be exceeded,
    /// or [`PatchError::Allocation`] when retaining the new payload fails.
    pub fn insert(&mut self, input: impl AsRef<[u8]>) -> Result<BlobId, PatchError> {
        let source = input.as_ref();
        self.ensure_blob_size(source.len())?;
        let id = BlobId::of(source);
        if self.entries.contains_key(&id) {
            return Ok(id);
        }
        self.ensure_room(source.len())?;
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(source.len())
            .map_err(|_allocation_error| PatchError::Allocation)?;
        owned.extend_from_slice(source);
        self.bytes = self
            .bytes
            .checked_add(source.len())
            .ok_or(PatchError::BlobLimit {
                kind: BlobLimitKind::TotalBytes,
                observed: usize::MAX,
                limit: self.limits.total_bytes,
            })?;
        self.entries.insert(id.clone(), Arc::from(owned));
        Ok(id)
    }

    /// Returns one payload by its content address.
    #[must_use]
    pub fn get(&self, id: &BlobId) -> Option<&[u8]> {
        self.entries.get(id).map(AsRef::as_ref)
    }

    /// Returns the bundle's configured bounds.
    #[must_use]
    pub const fn limits(&self) -> BlobLimits {
        self.limits
    }

    /// Number of distinct payloads.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Aggregate decoded payload byte length.
    #[must_use]
    pub const fn bytes_len(&self) -> usize {
        self.bytes
    }

    /// Whether no blobs are attached.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Content addresses in stable lexicographic order.
    #[must_use]
    pub fn ids(&self) -> impl ExactSizeIterator<Item = &BlobId> {
        self.entries.keys()
    }

    fn ensure_room(&self, length: usize) -> Result<(), PatchError> {
        if self.entries.len() >= self.limits.blobs {
            return Err(PatchError::BlobLimit {
                kind: BlobLimitKind::Count,
                observed: self.entries.len().saturating_add(1),
                limit: self.limits.blobs,
            });
        }
        let total = self
            .bytes
            .checked_add(length)
            .ok_or(PatchError::BlobLimit {
                kind: BlobLimitKind::TotalBytes,
                observed: usize::MAX,
                limit: self.limits.total_bytes,
            })?;
        if total > self.limits.total_bytes {
            return Err(PatchError::BlobLimit {
                kind: BlobLimitKind::TotalBytes,
                observed: total,
                limit: self.limits.total_bytes,
            });
        }
        Ok(())
    }

    fn ensure_blob_size(&self, length: usize) -> Result<(), PatchError> {
        if length > self.limits.blob_bytes {
            return Err(PatchError::BlobLimit {
                kind: BlobLimitKind::BlobBytes,
                observed: length,
                limit: self.limits.blob_bytes,
            });
        }
        Ok(())
    }
}

impl fmt::Debug for BlobBundle {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BlobBundle")
            .field("limits", &self.limits)
            .field("bytes", &self.bytes)
            .field("ids", &self.entries.keys().collect::<Vec<_>>())
            .finish()
    }
}

/// One bounded blob-bundle limit that rejected an operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BlobLimitKind {
    /// The number of distinct blob payloads.
    Count,
    /// The decoded byte length of one blob payload.
    BlobBytes,
    /// The aggregate decoded byte length of all blobs.
    TotalBytes,
}

/// A semantic operation. Formats define the operation vocabulary and target
/// selectors; this common representation carries it durably.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[allow(
    clippy::module_name_repetitions,
    reason = "the type is re-exported as a format-neutral public patch operation"
)]
pub struct PatchOperation {
    /// Format-owned operation name, such as `cell.set`.
    pub op: String,
    /// Format-owned semantic selector, never a private on-disk identifier.
    pub target: String,
    /// Expected values or hashes required before application, sorted by name.
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub preconditions: BTreeMap<String, Value>,
    /// Format-owned operation payload.
    pub value: Value,
}

impl PatchOperation {
    /// Creates and validates a semantic operation.
    ///
    /// # Errors
    ///
    /// Returns [`PatchError::InvalidText`] when the operation vocabulary or
    /// precondition names cannot be represented in the bounded wire format.
    pub fn new(
        limits: PatchLimits,
        op: impl Into<String>,
        target: impl Into<String>,
        preconditions: BTreeMap<String, Value>,
        value: Value,
    ) -> Result<Self, PatchError> {
        let operation = Self {
            op: op.into(),
            target: target.into(),
            preconditions,
            value,
        };
        validate_operation(&operation, limits)?;
        Ok(operation)
    }

    fn validate(&self) -> Result<(), PatchError> {
        validate_text("operation name", &self.op, MAX_OPERATION_NAME_BYTES, false)?;
        validate_text("operation target", &self.target, MAX_TARGET_BYTES, true)?;
        for key in self.preconditions.keys() {
            validate_text("precondition name", key, MAX_OPERATION_NAME_BYTES, false)?;
        }
        Ok(())
    }
}

/// A forward semantic operation paired with its exact semantic inverse.
#[derive(Debug, Clone, PartialEq)]
pub struct ReversibleOperation {
    forward: PatchOperation,
    inverse: PatchOperation,
}

impl ReversibleOperation {
    /// Creates one reversible operation pair.
    #[must_use]
    pub fn new(forward: PatchOperation, inverse: PatchOperation) -> Self {
        Self { forward, inverse }
    }

    /// The operation applied in the forward direction.
    #[must_use]
    pub const fn forward(&self) -> &PatchOperation {
        &self.forward
    }

    /// The operation applied in the reverse direction.
    #[must_use]
    pub const fn inverse(&self) -> &PatchOperation {
        &self.inverse
    }
}

/// A format-neutral, versioned durable patch.
///
/// `Patch<Reversible>` retains both directions and can be inverted or sealed.
/// `Patch<ForwardOnly>` has discarded all reverse semantic operations and
/// reverse-only blobs; it cannot be made reversible again.
#[derive(Clone)]
pub struct Patch<Mode> {
    format: String,
    version: u16,
    limits: PatchLimits,
    forward: Vec<PatchOperation>,
    forward_blobs: BlobBundle,
    reverse: Vec<PatchOperation>,
    reverse_blobs: BlobBundle,
    marker: PhantomData<Mode>,
}

impl Patch<Reversible> {
    /// Creates a reversible patch from format-owned semantic operation pairs.
    ///
    /// # Errors
    ///
    /// Returns [`PatchError::InvalidText`] when the format or an operation
    /// cannot be represented by the bounded wire vocabulary.
    pub fn new(
        limits: PatchLimits,
        format_name: impl Into<String>,
        operations: impl IntoIterator<Item = ReversibleOperation>,
        forward_blobs: BlobBundle,
        reverse_blobs: BlobBundle,
    ) -> Result<Self, PatchError> {
        let format = format_name.into();
        validate_text("format", &format, MAX_FORMAT_NAME_BYTES, false)?;
        validate_bundle_limits(&forward_blobs, limits.blobs())?;
        validate_bundle_limits(&reverse_blobs, limits.blobs())?;
        let mut forward = Vec::new();
        let mut reverse = Vec::new();
        for operation in operations {
            operation.forward.validate()?;
            operation.inverse.validate()?;
            forward.push(operation.forward);
            reverse.push(operation.inverse);
        }
        validate_reversible_operations(&forward, &reverse, limits)?;
        Ok(Self {
            format,
            version: PATCH_WIRE_VERSION,
            limits,
            forward,
            forward_blobs,
            reverse,
            reverse_blobs,
            marker: PhantomData,
        })
    }

    /// Produces the reverse patch without retaining redundant payload copies.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            format: self.format.clone(),
            version: self.version,
            limits: self.limits,
            forward: self.reverse.iter().rev().cloned().collect(),
            forward_blobs: self.reverse_blobs.clone(),
            reverse: self.forward.iter().rev().cloned().collect(),
            reverse_blobs: self.forward_blobs.clone(),
            marker: PhantomData,
        }
    }

    /// Consumes this patch and permanently discards all reverse material.
    #[must_use]
    pub fn seal(self) -> Patch<ForwardOnly> {
        let empty_reverse = BlobBundle::new(self.reverse_blobs.limits());
        Patch {
            format: self.format,
            version: self.version,
            limits: self.limits,
            forward: self.forward,
            forward_blobs: self.forward_blobs,
            reverse: Vec::new(),
            reverse_blobs: empty_reverse,
            marker: PhantomData,
        }
    }

    /// Serializes this patch in canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns [`PatchError::Json`] if a supplied semantic value cannot be
    /// represented as JSON.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>, PatchError> {
        preflight_reversible_json(self)?;
        let wire = ReversibleWire::from_patch(self);
        canonical_json(&wire, self.limits.json_bytes)
    }

    /// Parses only canonical deterministic JSON under explicit blob limits.
    ///
    /// # Errors
    ///
    /// Returns a parse, canonicality, integrity, or bound error when `bytes`
    /// is not an accepted durable reversible patch.
    pub fn from_deterministic_json(bytes: &[u8], limits: PatchLimits) -> Result<Self, PatchError> {
        let wire: ReversibleWire = parse_canonical(bytes, limits)?;
        wire.into_patch(limits)
    }

    /// Fingerprints the complete canonical reversible wire envelope.
    ///
    /// # Errors
    ///
    /// Returns the same bounded serialization error as
    /// [`Self::to_deterministic_json`].
    pub fn fingerprint(&self) -> Result<DiagnosticFingerprint, PatchError> {
        self.to_deterministic_json()
            .map(|wire| DiagnosticFingerprint::of(&wire))
    }
}

impl Patch<ForwardOnly> {
    /// Creates a forward-only patch. This is intended for formats that cannot
    /// safely retain inverse data in the first place.
    ///
    /// # Errors
    ///
    /// Returns [`PatchError::InvalidText`] when the format or an operation
    /// cannot be represented by the bounded wire vocabulary.
    pub fn new(
        limits: PatchLimits,
        format_name: impl Into<String>,
        operations: impl IntoIterator<Item = PatchOperation>,
        blobs: BlobBundle,
    ) -> Result<Self, PatchError> {
        let format = format_name.into();
        validate_text("format", &format, MAX_FORMAT_NAME_BYTES, false)?;
        let forward = operations.into_iter().collect::<Vec<_>>();
        validate_bundle_limits(&blobs, limits.blobs())?;
        validate_operations(&forward, limits)?;
        let empty_reverse = BlobBundle::new(blobs.limits());
        Ok(Self {
            format,
            version: PATCH_WIRE_VERSION,
            limits,
            forward,
            forward_blobs: blobs,
            reverse: Vec::new(),
            reverse_blobs: empty_reverse,
            marker: PhantomData,
        })
    }

    /// Serializes this patch in canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns [`PatchError::Json`] if a supplied semantic value cannot be
    /// represented as JSON.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>, PatchError> {
        preflight_forward_json(self)?;
        let wire = ForwardWire::from_patch(self);
        canonical_json(&wire, self.limits.json_bytes)
    }

    /// Parses only canonical deterministic JSON under explicit blob limits.
    ///
    /// # Errors
    ///
    /// Returns a parse, canonicality, integrity, or bound error when `bytes`
    /// is not an accepted durable forward-only patch.
    pub fn from_deterministic_json(bytes: &[u8], limits: PatchLimits) -> Result<Self, PatchError> {
        let wire: ForwardWire = parse_canonical(bytes, limits)?;
        wire.into_patch(limits)
    }

    /// Fingerprints the complete canonical forward-only wire envelope.
    ///
    /// # Errors
    ///
    /// Returns the same bounded serialization error as
    /// [`Self::to_deterministic_json`].
    pub fn fingerprint(&self) -> Result<DiagnosticFingerprint, PatchError> {
        self.to_deterministic_json()
            .map(|wire| DiagnosticFingerprint::of(&wire))
    }
}

impl<Mode> Patch<Mode> {
    /// Format namespace identifying the semantic operation vocabulary.
    #[must_use]
    pub fn format(&self) -> &str {
        &self.format
    }

    /// Durable wire-schema version.
    #[must_use]
    pub const fn version(&self) -> u16 {
        self.version
    }

    /// Finite bounds retained with this patch.
    #[must_use]
    pub const fn limits(&self) -> PatchLimits {
        self.limits
    }

    /// Forward operations in caller-defined deterministic order.
    #[must_use]
    pub fn operations(&self) -> &[PatchOperation] {
        &self.forward
    }

    /// Forward content-addressed blob bundle.
    #[must_use]
    pub const fn blobs(&self) -> &BlobBundle {
        &self.forward_blobs
    }
}

impl<Mode> fmt::Debug for Patch<Mode> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("format", &self.format)
            .field("version", &self.version)
            .field("operations", &self.forward)
            .field("blobs", &self.forward_blobs)
            .finish_non_exhaustive()
    }
}

/// A generic structured conflict collection. Format owners supply the entry
/// type so their conflict details remain strongly typed.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ConflictSet<C> {
    conflicts: Box<[C]>,
}

impl<C> ConflictSet<C> {
    /// Creates a conflict set in the deterministic order chosen by its owner.
    #[must_use]
    pub fn new(conflicts: impl Into<Box<[C]>>) -> Self {
        Self {
            conflicts: conflicts.into(),
        }
    }

    /// Returns conflicts in their owner-defined deterministic order.
    #[must_use]
    pub fn conflicts(&self) -> &[C] {
        &self.conflicts
    }

    /// Number of conflicts.
    #[must_use]
    pub fn len(&self) -> usize {
        self.conflicts.len()
    }

    /// Whether the set contains no conflicts.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.conflicts.is_empty()
    }

    /// Consumes the set into its ordered entries.
    #[must_use]
    pub fn into_boxed_slice(self) -> Box<[C]> {
        self.conflicts
    }
}

/// Finite retention bounds for undo/redo history.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HistoryLimits {
    steps: usize,
    weight: u64,
}

impl HistoryLimits {
    /// Creates finite history limits.
    #[must_use]
    pub const fn new(max_steps: usize, max_weight: u64) -> Self {
        Self {
            steps: max_steps,
            weight: max_weight,
        }
    }

    /// Maximum retained undo or redo transitions.
    #[must_use]
    pub const fn max_steps(self) -> usize {
        self.steps
    }

    /// Maximum retained transition weight.
    #[must_use]
    pub const fn max_weight(self) -> u64 {
        self.weight
    }
}

struct HistoryStep<T> {
    state: T,
    weight: u64,
}

/// Explicit, bounded undo/redo state.
///
/// `T` is normally a format-owned immutable snapshot. The caller supplies the
/// retained patch/blob byte cost for every transition, keeping storage policy
/// explicit instead of estimating arbitrary snapshot sizes in common code.
pub struct History<T> {
    current: T,
    undo: VecDeque<HistoryStep<T>>,
    redo: VecDeque<HistoryStep<T>>,
    limits: HistoryLimits,
    retained_weight: u64,
}

impl<T> History<T> {
    /// Starts history at one current snapshot with no retained transitions.
    #[must_use]
    pub fn new(current: T, limits: HistoryLimits) -> Self {
        Self {
            current,
            undo: VecDeque::new(),
            redo: VecDeque::new(),
            limits,
            retained_weight: 0,
        }
    }

    /// The currently selected snapshot.
    #[must_use]
    pub const fn current(&self) -> &T {
        &self.current
    }

    /// Configured finite retention bounds.
    #[must_use]
    pub const fn limits(&self) -> HistoryLimits {
        self.limits
    }

    /// Aggregate transition weight currently retained by the undo stack.
    #[must_use]
    pub const fn retained_weight(&self) -> u64 {
        self.retained_weight
    }

    /// Whether an undo transition is available.
    #[must_use]
    pub fn can_undo(&self) -> bool {
        !self.undo.is_empty()
    }

    /// Whether a redo transition is available.
    #[must_use]
    pub fn can_redo(&self) -> bool {
        !self.redo.is_empty()
    }

    /// Records a newly committed state and clears invalidated redo history.
    ///
    /// Oldest undo transitions are returned when explicit bounds require their
    /// eviction. A transition heavier than the entire budget is refused before
    /// changing the current state or either history stack.
    ///
    /// # Errors
    ///
    /// Returns [`PatchError::HistoryWeight`] without changing history when the
    /// new transition alone exceeds the configured retention budget.
    pub fn record(&mut self, next: T, weight: u64) -> Result<Vec<T>, PatchError> {
        if weight > self.limits.weight {
            return Err(PatchError::HistoryWeight {
                observed: weight,
                limit: self.limits.weight,
            });
        }

        let mut discarded = self
            .redo
            .drain(..)
            .map(|step| step.state)
            .collect::<Vec<_>>();
        while self.undo.len() >= self.limits.steps
            || self.retained_weight.saturating_add(weight) > self.limits.weight
        {
            let Some(oldest) = self.undo.front() else {
                break;
            };
            self.retained_weight = self.retained_weight.saturating_sub(oldest.weight);
            let evicted = self.undo.pop_front().ok_or(PatchError::HistoryInvariant)?;
            discarded.push(evicted.state);
        }
        if self.limits.steps == 0 {
            discarded.push(std::mem::replace(&mut self.current, next));
            return Ok(discarded);
        }
        self.undo.push_back(HistoryStep {
            state: std::mem::replace(&mut self.current, next),
            weight,
        });
        self.retained_weight = self.retained_weight.saturating_add(weight);
        Ok(discarded)
    }

    /// Moves one retained transition backward, returning whether it existed.
    pub fn undo(&mut self) -> bool {
        let Some(step) = self.undo.pop_back() else {
            return false;
        };
        self.retained_weight = self.retained_weight.saturating_sub(step.weight);
        self.redo.push_back(HistoryStep {
            state: std::mem::replace(&mut self.current, step.state),
            weight: step.weight,
        });
        true
    }

    /// Moves one retained transition forward, returning whether it existed.
    pub fn redo(&mut self) -> bool {
        let Some(step) = self.redo.pop_back() else {
            return false;
        };
        self.retained_weight = self.retained_weight.saturating_add(step.weight);
        self.undo.push_back(HistoryStep {
            state: std::mem::replace(&mut self.current, step.state),
            weight: step.weight,
        });
        true
    }
}

/// Failures while constructing, decoding, or retaining a durable patch.
#[derive(Debug, Error)]
#[allow(
    clippy::module_name_repetitions,
    reason = "the error is re-exported separately from litchi_core::Error"
)]
#[non_exhaustive]
pub enum PatchError {
    /// A required semantic identifier was empty, contained a control code, or
    /// exceeded its finite wire bound.
    #[error("invalid {field}")]
    InvalidText { field: &'static str },
    /// A blob count or decoded byte bound was exceeded.
    #[error("{kind:?} blob limit exceeded: observed {observed}, limit {limit}")]
    BlobLimit {
        /// Bound that failed.
        kind: BlobLimitKind,
        /// Requested or observed amount.
        observed: usize,
        /// Configured maximum.
        limit: usize,
    },
    /// A blob bundle was prepared under bounds different from its patch.
    #[error("blob bundle limits do not match patch limits")]
    IncompatibleBlobLimits {
        /// Bounds retained by the patch.
        expected: BlobLimits,
        /// Bounds used when the bundle was created.
        actual: BlobLimits,
    },
    /// Reversible forward and inverse operation lists had different lengths.
    #[error("reversible patch operation pairs are not aligned")]
    OperationPairing,
    /// A SHA-256 blob identifier was malformed or non-canonical.
    #[error("invalid blob SHA-256 identifier")]
    InvalidBlobId,
    /// A blob's declared content address did not match its decoded bytes.
    #[error("blob bytes do not match their declared SHA-256 identifier")]
    BlobDigestMismatch,
    /// Base64 decoding failed for an attached blob.
    #[error("invalid blob base64: {0}")]
    BlobBase64(#[source] base64::DecodeError),
    /// JSON parsing failed.
    #[error("invalid patch JSON: {0}")]
    Json(#[source] serde_json::Error),
    /// JSON was valid but did not use the canonical deterministic encoding.
    #[error("patch JSON is not canonical deterministic JSON")]
    NonCanonicalJson,
    /// The wire schema revision is not supported by this library version.
    #[error("unsupported patch wire version {0}")]
    UnsupportedVersion(u16),
    /// A history transition exceeds the full configured history budget.
    #[error("history weight {observed} exceeds limit {limit}")]
    HistoryWeight { observed: u64, limit: u64 },
    /// An internal history-stack invariant was violated.
    #[error("history stack invariant failed")]
    HistoryInvariant,
    /// An untrusted JSON input exceeded one finite wire bound.
    #[error("{kind:?} JSON limit exceeded: observed {observed}, limit {limit}")]
    JsonLimit {
        /// Bound that failed.
        kind: JsonLimitKind,
        /// Requested or observed amount.
        observed: usize,
        /// Configured maximum.
        limit: usize,
    },
    /// Allocating a bounded patch-owned byte buffer failed.
    #[error("allocation failed while retaining patch bytes")]
    Allocation,
}

/// One finite JSON wire bound that rejected an untrusted patch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum JsonLimitKind {
    /// Raw JSON byte length before parsing.
    InputBytes,
    /// Canonical JSON byte length while serializing.
    OutputBytes,
    /// Number of semantic operations.
    Operations,
    /// Nesting depth of a JSON semantic value.
    ValueDepth,
    /// UTF-8 byte length of one string in a semantic value.
    ValueStringBytes,
    /// Aggregate encoded byte length of semantic operations.
    PayloadBytes,
}

/// Explicit finite bounds for generic sub-edit composition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CompositionLimits {
    sub_edits: usize,
    effects_per_sub_edit: usize,
    total_effects: usize,
    conflicts: usize,
}

impl CompositionLimits {
    /// Creates finite bounds for one joined edit or merge plan.
    #[must_use]
    pub const fn new(
        max_sub_edits: usize,
        max_effects_per_sub_edit: usize,
        max_total_effects: usize,
        max_conflicts: usize,
    ) -> Self {
        Self {
            sub_edits: max_sub_edits,
            effects_per_sub_edit: max_effects_per_sub_edit,
            total_effects: max_total_effects,
            conflicts: max_conflicts,
        }
    }

    /// Maximum sub-edits retained by one composition.
    #[must_use]
    pub const fn max_sub_edits(self) -> usize {
        self.sub_edits
    }

    /// Maximum declared effects accepted from one sub-edit.
    #[must_use]
    pub const fn max_effects_per_sub_edit(self) -> usize {
        self.effects_per_sub_edit
    }

    /// Maximum canonical effects retained by one composition.
    #[must_use]
    pub const fn max_total_effects(self) -> usize {
        self.total_effects
    }

    /// Maximum detailed conflicts returned by one operation.
    #[must_use]
    pub const fn max_conflicts(self) -> usize {
        self.conflicts
    }
}

/// One finite sub-edit composition bound.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum CompositionLimitKind {
    /// Number of sub-edits.
    SubEdits,
    /// Effects declared by one sub-edit before deduplication.
    EffectsPerSubEdit,
    /// Aggregate canonical effects.
    TotalEffects,
    /// Detailed conflict entries.
    Conflicts,
}

/// Failure while constructing or planning bounded sub-edits.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum CompositionError {
    /// A stable sub-edit identifier was empty, contained a control code, or
    /// exceeded the fixed durable operation-name bound.
    #[error("invalid sub-edit identifier")]
    InvalidId,
    /// An exact semantic effect key was empty, contained a control code, or
    /// exceeded the fixed durable target bound.
    #[error("invalid semantic effect key")]
    InvalidEffect,
    /// One finite composition bound was exceeded.
    #[error("{kind:?} composition limit exceeded: observed {observed}, limit {limit}")]
    Limit {
        /// Bound that failed.
        kind: CompositionLimitKind,
        /// Requested or observed amount.
        observed: usize,
        /// Configured maximum.
        limit: usize,
    },
}

/// Whether a sub-edit observes or changes one exact semantic effect key.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
#[non_exhaustive]
pub enum EffectAccess {
    /// The staged operation depends on current semantic state.
    Read,
    /// The staged operation changes semantic state.
    Write,
}

/// Canonical exact-key read and write effects for one sub-edit.
///
/// Keys are deliberately opaque to common code. Format owners map ranges,
/// structural dependencies, and effect facets that can interfere to an equal
/// key; this layer never guesses hierarchy from string prefixes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EffectSet {
    reads: BTreeSet<String>,
    writes: BTreeSet<String>,
}

impl EffectSet {
    /// Exact keys read but not written by this sub-edit.
    #[must_use]
    pub fn reads(&self) -> impl ExactSizeIterator<Item = &str> {
        self.reads.iter().map(String::as_str)
    }

    /// Exact keys written by this sub-edit.
    #[must_use]
    pub fn writes(&self) -> impl ExactSizeIterator<Item = &str> {
        self.writes.iter().map(String::as_str)
    }

    /// Number of canonical read and write effects.
    #[must_use]
    pub fn len(&self) -> usize {
        self.reads.len().saturating_add(self.writes.len())
    }

    /// Whether no effect is declared.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.reads.is_empty() && self.writes.is_empty()
    }
}

/// One exact-key read/write overlap between two sub-edits.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
pub struct EffectConflict {
    effect: String,
    left_id: String,
    right_id: String,
    left_access: EffectAccess,
    right_access: EffectAccess,
}

impl EffectConflict {
    /// Exact format-owned semantic key that overlaps.
    #[must_use]
    pub fn effect(&self) -> &str {
        &self.effect
    }

    /// Stable left sub-edit identifier.
    #[must_use]
    pub fn left_id(&self) -> &str {
        &self.left_id
    }

    /// Stable right sub-edit identifier.
    #[must_use]
    pub fn right_id(&self) -> &str {
        &self.right_id
    }

    /// Left access participating in the overlap.
    #[must_use]
    pub const fn left_access(&self) -> EffectAccess {
        self.left_access
    }

    /// Right access participating in the overlap.
    #[must_use]
    pub const fn right_access(&self) -> EffectAccess {
        self.right_access
    }
}

/// One reason two sub-edits cannot be combined automatically.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
#[non_exhaustive]
pub enum SubEditConflict {
    /// Both branches used the same stable sub-edit identifier.
    DuplicateId(String),
    /// At least one side writes an exact semantic key used by the other.
    Effect(EffectConflict),
}

/// Independently prepared staged work against one exact immutable base.
///
/// `L` is a format-owned exact source-lineage token, not a diagnostic hash.
/// `T` is the format-owned staged operation payload.
pub struct SubEdit<L, T> {
    lineage: L,
    limits: CompositionLimits,
    id: String,
    effects: EffectSet,
    payload: T,
}

impl<L, T> SubEdit<L, T> {
    /// Creates one bounded sub-edit and canonicalizes its effects.
    ///
    /// A key declared as both read and written is retained only as a write,
    /// while all input occurrences are charged before deduplication.
    ///
    /// # Errors
    ///
    /// Returns a text or per-sub-edit effect-bound error.
    pub fn new(
        lineage: L,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        reads: impl IntoIterator<Item = String>,
        writes: impl IntoIterator<Item = String>,
        payload: T,
    ) -> Result<Self, CompositionError> {
        let id = identifier.into();
        if !valid_composition_text(&id, MAX_OPERATION_NAME_BYTES) {
            return Err(CompositionError::InvalidId);
        }
        let effects = collect_effect_set(reads, writes, limits)?;
        Ok(Self {
            lineage,
            limits,
            id,
            effects,
            payload,
        })
    }

    /// Exact format-owned source-lineage token.
    #[must_use]
    pub const fn lineage(&self) -> &L {
        &self.lineage
    }

    /// Stable composition identifier.
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Canonical read and write effects.
    #[must_use]
    pub const fn effects(&self) -> &EffectSet {
        &self.effects
    }

    /// Borrow the format-owned staged payload.
    #[must_use]
    pub const fn payload(&self) -> &T {
        &self.payload
    }

    /// Consumes the wrapper and returns the staged payload.
    #[must_use]
    pub fn into_payload(self) -> T {
        self.payload
    }
}

impl<L, T> fmt::Debug for SubEdit<L, T> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SubEdit")
            .field("id", &self.id)
            .field("effects", &self.effects)
            .finish_non_exhaustive()
    }
}

/// A bounded deterministic collection of provably disjoint sub-edits.
pub struct JoinedSubEdits<L, T> {
    lineage: L,
    limits: CompositionLimits,
    total_effects: usize,
    edits: BTreeMap<String, SubEdit<L, T>>,
}

impl<L, T> JoinedSubEdits<L, T> {
    /// Starts an empty composition for one exact source lineage.
    #[must_use]
    pub fn new(lineage: L, limits: CompositionLimits) -> Self {
        Self {
            lineage,
            limits,
            total_effects: 0,
            edits: BTreeMap::new(),
        }
    }

    /// Exact format-owned source-lineage token.
    #[must_use]
    pub const fn lineage(&self) -> &L {
        &self.lineage
    }

    /// Accepted sub-edits in stable identifier order.
    #[must_use]
    pub fn sub_edits(&self) -> impl ExactSizeIterator<Item = &SubEdit<L, T>> {
        self.edits.values()
    }

    /// Number of accepted sub-edits.
    #[must_use]
    pub fn len(&self) -> usize {
        self.edits.len()
    }

    /// Whether no sub-edit has been accepted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.edits.is_empty()
    }

    /// Aggregate number of canonical effects.
    #[must_use]
    pub const fn total_effects(&self) -> usize {
        self.total_effects
    }

    /// Consumes the composition into stable identifier order.
    #[must_use]
    pub fn into_sub_edits(self) -> impl ExactSizeIterator<Item = SubEdit<L, T>> {
        self.edits.into_values()
    }

    fn insert_disjoint(&mut self, edit: SubEdit<L, T>) {
        self.total_effects = self.total_effects.saturating_add(edit.effects.len());
        self.edits.insert(edit.id.clone(), edit);
    }
}

impl<L: Eq, T> JoinedSubEdits<L, T> {
    /// Joins one sub-edit only when its exact effects are provably disjoint.
    ///
    /// On failure `self` remains unchanged and the error returns `incoming`.
    ///
    /// # Errors
    ///
    /// Returns a lineage, limits, identifier, overlap, or finite-bound error.
    pub fn join(&mut self, incoming: SubEdit<L, T>) -> Result<&mut Self, SubEditJoinError<L, T>> {
        if let Some(failure) = join_failure(self, &incoming) {
            return Err(SubEditJoinError {
                failure,
                rejected: Box::new(incoming),
            });
        }
        self.insert_disjoint(incoming);
        Ok(self)
    }
}

impl<L, T> fmt::Debug for JoinedSubEdits<L, T> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("JoinedSubEdits")
            .field("limits", &self.limits)
            .field("total_effects", &self.total_effects)
            .field("ids", &self.edits.keys().collect::<Vec<_>>())
            .finish_non_exhaustive()
    }
}

/// Why an independently prepared sub-edit could not be joined.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum SubEditJoinFailure {
    /// Exact source lineage differs.
    DifferentLineage,
    /// The incoming sub-edit uses different finite bounds.
    DifferentLimits,
    /// The stable identifier already exists.
    DuplicateId,
    /// One or more exact effects overlap.
    Overlap(ConflictSet<SubEditConflict>),
    /// Accepting or fully reporting the edit would exceed a bound.
    Limit(CompositionError),
}

/// Recoverable join failure retaining the rejected sub-edit.
pub struct SubEditJoinError<L, T> {
    failure: SubEditJoinFailure,
    rejected: Box<SubEdit<L, T>>,
}

impl<L, T> SubEditJoinError<L, T> {
    /// Structured refusal reason.
    #[must_use]
    pub const fn failure(&self) -> &SubEditJoinFailure {
        &self.failure
    }

    /// Borrow the rejected sub-edit.
    #[must_use]
    pub const fn rejected(&self) -> &SubEdit<L, T> {
        &self.rejected
    }

    /// Recover the rejected sub-edit.
    #[must_use]
    pub fn into_rejected(self) -> SubEdit<L, T> {
        *self.rejected
    }

    /// Recover both the reason and rejected sub-edit.
    #[must_use]
    pub fn into_parts(self) -> (SubEditJoinFailure, SubEdit<L, T>) {
        (self.failure, *self.rejected)
    }
}

impl<L, T> fmt::Debug for SubEditJoinError<L, T> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SubEditJoinError")
            .field("failure", &self.failure)
            .field("rejected_id", &self.rejected.id)
            .finish()
    }
}

/// Explicit resolution of the conflicting portion of a three-way merge.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum MergeChoice {
    /// Keep conflicting sub-edits from the left branch.
    Left,
    /// Keep conflicting sub-edits from the right branch.
    Right,
    /// Drop conflicting sub-edits from both branches.
    Neither,
}

/// A non-applying three-way plan for two branches from one exact base.
///
/// Disjoint sub-edits from both branches are retained automatically. All
/// overlapping sub-edits form one conservative conflict group and require an
/// explicit choice before [`Self::finish`] yields staged work. A format owner
/// still validates and commits that work atomically against its exact base.
pub struct ThreeWayMergePlan<L, T> {
    automatic: JoinedSubEdits<L, T>,
    left_conflicts: JoinedSubEdits<L, T>,
    right_conflicts: JoinedSubEdits<L, T>,
    conflicts: ConflictSet<SubEditConflict>,
    resolution: Option<MergeChoice>,
}

impl<L: Clone + Eq, T> ThreeWayMergePlan<L, T> {
    /// Plans a bounded merge without applying either branch.
    ///
    /// # Errors
    ///
    /// Returns both branches intact when lineage or limits differ, combined
    /// work exceeds a bound, or all conflicts cannot be reported within the
    /// configured conflict limit.
    pub fn new(
        left: JoinedSubEdits<L, T>,
        right: JoinedSubEdits<L, T>,
    ) -> Result<Self, ThreeWayMergeError<L, T>> {
        build_three_way_plan(left, right)
    }
}

impl<L, T> ThreeWayMergePlan<L, T> {
    /// Automatically accepted disjoint work from both branches.
    #[must_use]
    pub const fn automatic(&self) -> &JoinedSubEdits<L, T> {
        &self.automatic
    }

    /// Conflicting left-branch work in stable identifier order.
    #[must_use]
    pub const fn left_conflicts(&self) -> &JoinedSubEdits<L, T> {
        &self.left_conflicts
    }

    /// Conflicting right-branch work in stable identifier order.
    #[must_use]
    pub const fn right_conflicts(&self) -> &JoinedSubEdits<L, T> {
        &self.right_conflicts
    }

    /// Deterministically ordered overlap details.
    #[must_use]
    pub const fn conflicts(&self) -> &ConflictSet<SubEditConflict> {
        &self.conflicts
    }

    /// Current explicit resolution, if conflicts exist and were resolved.
    #[must_use]
    pub const fn resolution(&self) -> Option<MergeChoice> {
        self.resolution
    }

    /// Resolves the complete conservative conflict group.
    pub fn resolve(&mut self, choice: MergeChoice) -> &mut Self {
        self.resolution = Some(choice);
        self
    }

    /// Produces complete staged work only after every conflict is resolved.
    ///
    /// # Errors
    ///
    /// Returns this plan unchanged while conflicts remain unresolved.
    pub fn finish(mut self) -> Result<JoinedSubEdits<L, T>, Box<Self>> {
        if !self.conflicts.is_empty() && self.resolution.is_none() {
            return Err(Box::new(self));
        }
        let selected = match self.resolution {
            Some(MergeChoice::Left) => self.left_conflicts.edits,
            Some(MergeChoice::Right) => self.right_conflicts.edits,
            Some(MergeChoice::Neither) | None => BTreeMap::new(),
        };
        for edit in selected.into_values() {
            self.automatic.insert_disjoint(edit);
        }
        Ok(self.automatic)
    }
}

impl<L, T> fmt::Debug for ThreeWayMergePlan<L, T> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ThreeWayMergePlan")
            .field("automatic", &self.automatic)
            .field("left_conflicts", &self.left_conflicts)
            .field("right_conflicts", &self.right_conflicts)
            .field("conflicts", &self.conflicts)
            .field("resolution", &self.resolution)
            .finish()
    }
}

/// Why a three-way plan could not be constructed.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum ThreeWayMergeFailure {
    /// Exact format-owned base lineage differs.
    DifferentLineage,
    /// Branches use different finite bounds.
    DifferentLimits,
    /// Planning every possible resolution would exceed a finite bound.
    Limit(CompositionError),
}

/// Recoverable three-way planning failure retaining both branches.
pub struct ThreeWayMergeError<L, T> {
    failure: ThreeWayMergeFailure,
    left: Box<JoinedSubEdits<L, T>>,
    right: Box<JoinedSubEdits<L, T>>,
}

impl<L, T> ThreeWayMergeError<L, T> {
    /// Structured planning refusal reason.
    #[must_use]
    pub const fn failure(&self) -> &ThreeWayMergeFailure {
        &self.failure
    }

    /// Recovers both unchanged branches.
    #[must_use]
    pub fn into_branches(self) -> (JoinedSubEdits<L, T>, JoinedSubEdits<L, T>) {
        (*self.left, *self.right)
    }
}

impl<L, T> fmt::Debug for ThreeWayMergeError<L, T> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ThreeWayMergeError")
            .field("failure", &self.failure)
            .field("left", &self.left)
            .field("right", &self.right)
            .finish()
    }
}

struct BoundedJsonOutput {
    bytes: Vec<u8>,
    maximum: usize,
}

impl BoundedJsonOutput {
    fn new(maximum: usize) -> Self {
        Self {
            bytes: Vec::new(),
            maximum,
        }
    }

    fn into_inner(self) -> Vec<u8> {
        self.bytes
    }

    fn write(&mut self, bytes: &[u8]) -> Result<(), PatchError> {
        let observed = self
            .bytes
            .len()
            .checked_add(bytes.len())
            .ok_or(PatchError::JsonLimit {
                kind: JsonLimitKind::OutputBytes,
                observed: usize::MAX,
                limit: self.maximum,
            })?;
        if observed > self.maximum {
            return Err(PatchError::JsonLimit {
                kind: JsonLimitKind::OutputBytes,
                observed,
                limit: self.maximum,
            });
        }
        self.bytes
            .try_reserve_exact(bytes.len())
            .map_err(|_allocation_error| PatchError::Allocation)?;
        self.bytes.extend_from_slice(bytes);
        Ok(())
    }
}

#[derive(Serialize, Deserialize)]
struct ForwardWire {
    blobs: Vec<WireBlob>,
    format: String,
    operations: Vec<PatchOperation>,
    version: u16,
}

impl ForwardWire {
    fn from_patch(patch: &Patch<ForwardOnly>) -> Self {
        Self {
            blobs: wire_blobs(&patch.forward_blobs),
            format: patch.format.clone(),
            operations: patch.forward.clone(),
            version: patch.version,
        }
    }

    fn into_patch(self, limits: PatchLimits) -> Result<Patch<ForwardOnly>, PatchError> {
        validate_version(self.version)?;
        let blobs = decode_blobs(self.blobs, limits.blobs())?;
        let mut patch = Patch::<ForwardOnly>::new(limits, self.format, self.operations, blobs)?;
        patch.version = self.version;
        Ok(patch)
    }
}

#[derive(Serialize, Deserialize)]
struct ReversibleWire {
    format: String,
    forward_blobs: Vec<WireBlob>,
    operations: Vec<WireReversibleOperation>,
    reverse_blobs: Vec<WireBlob>,
    version: u16,
}

impl ReversibleWire {
    fn from_patch(patch: &Patch<Reversible>) -> Self {
        let operations = patch
            .forward
            .iter()
            .cloned()
            .zip(patch.reverse.iter().cloned())
            .map(|(forward, inverse)| WireReversibleOperation { forward, inverse })
            .collect();
        Self {
            format: patch.format.clone(),
            forward_blobs: wire_blobs(&patch.forward_blobs),
            operations,
            reverse_blobs: wire_blobs(&patch.reverse_blobs),
            version: patch.version,
        }
    }

    fn into_patch(self, limits: PatchLimits) -> Result<Patch<Reversible>, PatchError> {
        validate_version(self.version)?;
        let forward_blobs = decode_blobs(self.forward_blobs, limits.blobs())?;
        let reverse_blobs = decode_blobs(self.reverse_blobs, limits.blobs())?;
        let operations = self
            .operations
            .into_iter()
            .map(|operation| ReversibleOperation::new(operation.forward, operation.inverse));
        let mut patch = Patch::<Reversible>::new(
            limits,
            self.format,
            operations,
            forward_blobs,
            reverse_blobs,
        )?;
        patch.version = self.version;
        Ok(patch)
    }
}

#[derive(Serialize, Deserialize)]
struct WireReversibleOperation {
    forward: PatchOperation,
    inverse: PatchOperation,
}

#[derive(Serialize, Deserialize)]
struct WireBlob {
    bytes: String,
    sha256: String,
}

type BranchConflictDetails = (Vec<SubEditConflict>, BTreeSet<String>, BTreeSet<String>);

fn valid_composition_text(text: &str, maximum: usize) -> bool {
    !text.is_empty() && text.len() <= maximum && !text.chars().any(char::is_control)
}

fn collect_effect_set(
    reads: impl IntoIterator<Item = String>,
    writes: impl IntoIterator<Item = String>,
    limits: CompositionLimits,
) -> Result<EffectSet, CompositionError> {
    let mut observed = 0_usize;
    let mut read_set = BTreeSet::new();
    for effect in reads {
        observed = observed.saturating_add(1);
        validate_effect_input(&effect, observed, limits)?;
        read_set.insert(effect);
    }
    let mut write_set = BTreeSet::new();
    for effect in writes {
        observed = observed.saturating_add(1);
        validate_effect_input(&effect, observed, limits)?;
        write_set.insert(effect);
    }
    read_set.retain(|effect| !write_set.contains(effect));
    Ok(EffectSet {
        reads: read_set,
        writes: write_set,
    })
}

fn validate_effect_input(
    effect: &str,
    observed: usize,
    limits: CompositionLimits,
) -> Result<(), CompositionError> {
    if observed > limits.effects_per_sub_edit {
        return Err(CompositionError::Limit {
            kind: CompositionLimitKind::EffectsPerSubEdit,
            observed,
            limit: limits.effects_per_sub_edit,
        });
    }
    if !valid_composition_text(effect, MAX_TARGET_BYTES) {
        return Err(CompositionError::InvalidEffect);
    }
    Ok(())
}

fn join_failure<L: Eq, T>(
    accepted: &JoinedSubEdits<L, T>,
    incoming: &SubEdit<L, T>,
) -> Option<SubEditJoinFailure> {
    if accepted.lineage != incoming.lineage {
        return Some(SubEditJoinFailure::DifferentLineage);
    }
    if accepted.limits != incoming.limits {
        return Some(SubEditJoinFailure::DifferentLimits);
    }
    if accepted.edits.contains_key(&incoming.id) {
        return Some(SubEditJoinFailure::DuplicateId);
    }
    if let Err(error) = ensure_composition_capacity(
        accepted.edits.len(),
        accepted.total_effects,
        1,
        incoming.effects.len(),
        accepted.limits,
    ) {
        return Some(SubEditJoinFailure::Limit(error));
    }
    let mut conflicts = Vec::new();
    for edit in accepted.edits.values() {
        let remaining = accepted.limits.conflicts.saturating_sub(conflicts.len());
        match effect_conflicts(edit, incoming, remaining) {
            Ok(found) => conflicts.extend(found.into_iter().map(SubEditConflict::Effect)),
            Err(error) => return Some(SubEditJoinFailure::Limit(error)),
        }
    }
    if conflicts.is_empty() {
        None
    } else {
        conflicts.sort_unstable();
        Some(SubEditJoinFailure::Overlap(ConflictSet::new(conflicts)))
    }
}

fn build_three_way_plan<L: Clone + Eq, T>(
    left: JoinedSubEdits<L, T>,
    right: JoinedSubEdits<L, T>,
) -> Result<ThreeWayMergePlan<L, T>, ThreeWayMergeError<L, T>> {
    if left.lineage != right.lineage {
        return Err(three_way_error(
            ThreeWayMergeFailure::DifferentLineage,
            left,
            right,
        ));
    }
    if left.limits != right.limits {
        return Err(three_way_error(
            ThreeWayMergeFailure::DifferentLimits,
            left,
            right,
        ));
    }
    if let Err(error) = ensure_composition_capacity(
        left.edits.len(),
        left.total_effects,
        right.edits.len(),
        right.total_effects,
        left.limits,
    ) {
        return Err(three_way_error(
            ThreeWayMergeFailure::Limit(error),
            left,
            right,
        ));
    }
    let (details, left_ids, right_ids) = match branch_conflicts(&left, &right) {
        Ok(result) => result,
        Err(error) => {
            return Err(three_way_error(
                ThreeWayMergeFailure::Limit(error),
                left,
                right,
            ));
        },
    };

    let JoinedSubEdits {
        lineage,
        limits,
        edits: left_edits,
        ..
    } = left;
    let JoinedSubEdits {
        edits: right_edits, ..
    } = right;
    let mut automatic = JoinedSubEdits::new(lineage.clone(), limits);
    let mut left_conflicts = JoinedSubEdits::new(lineage.clone(), limits);
    let mut right_conflicts = JoinedSubEdits::new(lineage, limits);
    partition_branch(left_edits, &left_ids, &mut automatic, &mut left_conflicts);
    partition_branch(
        right_edits,
        &right_ids,
        &mut automatic,
        &mut right_conflicts,
    );
    Ok(ThreeWayMergePlan {
        automatic,
        left_conflicts,
        right_conflicts,
        conflicts: ConflictSet::new(details),
        resolution: None,
    })
}

fn ensure_composition_capacity(
    left_edits: usize,
    left_effects: usize,
    right_edits: usize,
    right_effects: usize,
    limits: CompositionLimits,
) -> Result<(), CompositionError> {
    let edits = left_edits.saturating_add(right_edits);
    if edits > limits.sub_edits {
        return Err(CompositionError::Limit {
            kind: CompositionLimitKind::SubEdits,
            observed: edits,
            limit: limits.sub_edits,
        });
    }
    let effects = left_effects.saturating_add(right_effects);
    if effects > limits.total_effects {
        return Err(CompositionError::Limit {
            kind: CompositionLimitKind::TotalEffects,
            observed: effects,
            limit: limits.total_effects,
        });
    }
    Ok(())
}

fn effect_conflicts<L, T>(
    left: &SubEdit<L, T>,
    right: &SubEdit<L, T>,
    maximum: usize,
) -> Result<Vec<EffectConflict>, CompositionError> {
    let mut conflicts = Vec::new();
    append_effect_conflicts(
        &mut conflicts,
        left,
        right,
        left.effects.writes.intersection(&right.effects.writes),
        (EffectAccess::Write, EffectAccess::Write),
        maximum,
    )?;
    append_effect_conflicts(
        &mut conflicts,
        left,
        right,
        left.effects.writes.intersection(&right.effects.reads),
        (EffectAccess::Write, EffectAccess::Read),
        maximum,
    )?;
    append_effect_conflicts(
        &mut conflicts,
        left,
        right,
        left.effects.reads.intersection(&right.effects.writes),
        (EffectAccess::Read, EffectAccess::Write),
        maximum,
    )?;
    conflicts.sort_unstable();
    Ok(conflicts)
}

fn append_effect_conflicts<'a, L, T>(
    conflicts: &mut Vec<EffectConflict>,
    left: &SubEdit<L, T>,
    right: &SubEdit<L, T>,
    effects: impl Iterator<Item = &'a String>,
    access: (EffectAccess, EffectAccess),
    maximum: usize,
) -> Result<(), CompositionError> {
    for effect in effects {
        if conflicts.len() >= maximum {
            return Err(CompositionError::Limit {
                kind: CompositionLimitKind::Conflicts,
                observed: conflicts.len().saturating_add(1),
                limit: maximum,
            });
        }
        conflicts.push(EffectConflict {
            effect: effect.clone(),
            left_id: left.id.clone(),
            right_id: right.id.clone(),
            left_access: access.0,
            right_access: access.1,
        });
    }
    Ok(())
}

fn branch_conflicts<L, T>(
    left: &JoinedSubEdits<L, T>,
    right: &JoinedSubEdits<L, T>,
) -> Result<BranchConflictDetails, CompositionError> {
    let mut details = Vec::new();
    let mut left_ids = BTreeSet::new();
    let mut right_ids = BTreeSet::new();
    for left_edit in left.edits.values() {
        for right_edit in right.edits.values() {
            let mut pair = Vec::new();
            if left_edit.id == right_edit.id {
                pair.push(SubEditConflict::DuplicateId(left_edit.id.clone()));
            }
            let remaining = left
                .limits
                .conflicts
                .saturating_sub(details.len().saturating_add(pair.len()));
            pair.extend(
                effect_conflicts(left_edit, right_edit, remaining)?
                    .into_iter()
                    .map(SubEditConflict::Effect),
            );
            if pair.is_empty() {
                continue;
            }
            let observed = details.len().saturating_add(pair.len());
            if observed > left.limits.conflicts {
                return Err(CompositionError::Limit {
                    kind: CompositionLimitKind::Conflicts,
                    observed,
                    limit: left.limits.conflicts,
                });
            }
            left_ids.insert(left_edit.id.clone());
            right_ids.insert(right_edit.id.clone());
            details.extend(pair);
        }
    }
    details.sort_unstable();
    Ok((details, left_ids, right_ids))
}

fn partition_branch<L, T>(
    edits: BTreeMap<String, SubEdit<L, T>>,
    conflicting_ids: &BTreeSet<String>,
    automatic: &mut JoinedSubEdits<L, T>,
    conflicts: &mut JoinedSubEdits<L, T>,
) {
    for (id, edit) in edits {
        if conflicting_ids.contains(&id) {
            conflicts.insert_disjoint(edit);
        } else {
            automatic.insert_disjoint(edit);
        }
    }
}

fn three_way_error<L, T>(
    failure: ThreeWayMergeFailure,
    left: JoinedSubEdits<L, T>,
    right: JoinedSubEdits<L, T>,
) -> ThreeWayMergeError<L, T> {
    ThreeWayMergeError {
        failure,
        left: Box::new(left),
        right: Box::new(right),
    }
}

fn wire_blobs(bundle: &BlobBundle) -> Vec<WireBlob> {
    bundle
        .entries
        .iter()
        .map(|(id, bytes)| WireBlob {
            bytes: BASE64.encode(bytes),
            sha256: id.as_hex(),
        })
        .collect()
}

fn decode_blobs(blobs: Vec<WireBlob>, limits: BlobLimits) -> Result<BlobBundle, PatchError> {
    if blobs.len() > limits.blobs {
        return Err(PatchError::BlobLimit {
            kind: BlobLimitKind::Count,
            observed: blobs.len(),
            limit: limits.blobs,
        });
    }
    let mut bundle = BlobBundle::new(limits);
    let mut previous_id: Option<BlobId> = None;
    for blob in blobs {
        let maximum_encoded = base64_encoded_len(limits.blob_bytes)?;
        if blob.bytes.len() > maximum_encoded {
            return Err(PatchError::BlobLimit {
                kind: BlobLimitKind::BlobBytes,
                observed: base64_decoded_upper_bound(blob.bytes.len())?,
                limit: limits.blob_bytes,
            });
        }
        let id = BlobId::parse_hex(&blob.sha256)?;
        if previous_id.as_ref().is_some_and(|prior_id| prior_id >= &id) {
            return Err(PatchError::NonCanonicalJson);
        }
        let bytes = BASE64.decode(blob.bytes).map_err(PatchError::BlobBase64)?;
        if BlobId::of(&bytes) != id {
            return Err(PatchError::BlobDigestMismatch);
        }
        let inserted = bundle.insert(&bytes)?;
        if inserted != id {
            return Err(PatchError::BlobDigestMismatch);
        }
        previous_id = Some(id);
    }
    Ok(bundle)
}

fn base64_encoded_len(decoded: usize) -> Result<usize, PatchError> {
    decoded
        .checked_add(2)
        .and_then(|value| value.checked_div(3))
        .and_then(|value| value.checked_mul(4))
        .ok_or(PatchError::BlobLimit {
            kind: BlobLimitKind::BlobBytes,
            observed: usize::MAX,
            limit: usize::MAX,
        })
}

fn base64_decoded_upper_bound(encoded: usize) -> Result<usize, PatchError> {
    encoded
        .checked_add(3)
        .and_then(|value| value.checked_div(4))
        .and_then(|value| value.checked_mul(3))
        .ok_or(PatchError::BlobLimit {
            kind: BlobLimitKind::BlobBytes,
            observed: usize::MAX,
            limit: usize::MAX,
        })
}

fn validate_bundle_limits(bundle: &BlobBundle, expected: BlobLimits) -> Result<(), PatchError> {
    if bundle.limits != expected {
        return Err(PatchError::IncompatibleBlobLimits {
            expected,
            actual: bundle.limits,
        });
    }
    Ok(())
}

fn validate_operations(
    operations: &[PatchOperation],
    limits: PatchLimits,
) -> Result<(), PatchError> {
    if operations.len() > limits.operations {
        return Err(PatchError::JsonLimit {
            kind: JsonLimitKind::Operations,
            observed: operations.len(),
            limit: limits.operations,
        });
    }
    let mut payload_bytes = 0_usize;
    for operation in operations {
        payload_bytes = add_operation_payload(payload_bytes, operation, limits)?;
    }
    Ok(())
}

fn validate_reversible_operations(
    forward: &[PatchOperation],
    reverse: &[PatchOperation],
    limits: PatchLimits,
) -> Result<(), PatchError> {
    if forward.len() != reverse.len() {
        return Err(PatchError::OperationPairing);
    }
    if forward.len() > limits.operations {
        return Err(PatchError::JsonLimit {
            kind: JsonLimitKind::Operations,
            observed: forward.len(),
            limit: limits.operations,
        });
    }
    let mut payload_bytes = 0_usize;
    for operation in forward.iter().chain(reverse) {
        payload_bytes = add_operation_payload(payload_bytes, operation, limits)?;
    }
    Ok(())
}

fn add_operation_payload(
    payload_bytes: usize,
    operation: &PatchOperation,
    limits: PatchLimits,
) -> Result<usize, PatchError> {
    operation.validate()?;
    validate_semantic_value(&operation.value, limits)?;
    for value in operation.preconditions.values() {
        validate_semantic_value(value, limits)?;
    }
    let operation_bytes = operation_json_len(operation, limits.payload_bytes)?;
    let total = payload_bytes
        .checked_add(operation_bytes)
        .ok_or(PatchError::JsonLimit {
            kind: JsonLimitKind::PayloadBytes,
            observed: usize::MAX,
            limit: limits.payload_bytes,
        })?;
    if total > limits.payload_bytes {
        return Err(PatchError::JsonLimit {
            kind: JsonLimitKind::PayloadBytes,
            observed: total,
            limit: limits.payload_bytes,
        });
    }
    Ok(total)
}

fn validate_operation(operation: &PatchOperation, limits: PatchLimits) -> Result<(), PatchError> {
    let _payload_bytes = add_operation_payload(0, operation, limits)?;
    Ok(())
}

fn validate_semantic_value(value: &Value, limits: PatchLimits) -> Result<(), PatchError> {
    let mut stack = vec![(value, 1_usize)];
    while let Some((current, depth)) = stack.pop() {
        if depth > limits.value_depth {
            return Err(PatchError::JsonLimit {
                kind: JsonLimitKind::ValueDepth,
                observed: depth,
                limit: limits.value_depth,
            });
        }
        match current {
            Value::String(text) if text.len() > limits.value_string_bytes => {
                return Err(PatchError::JsonLimit {
                    kind: JsonLimitKind::ValueStringBytes,
                    observed: text.len(),
                    limit: limits.value_string_bytes,
                });
            },
            Value::Array(values) => {
                for child in values {
                    stack.push((child, depth.saturating_add(1)));
                }
            },
            Value::Object(values) => {
                for (key, child) in values {
                    if key.len() > limits.value_string_bytes {
                        return Err(PatchError::JsonLimit {
                            kind: JsonLimitKind::ValueStringBytes,
                            observed: key.len(),
                            limit: limits.value_string_bytes,
                        });
                    }
                    stack.push((child, depth.saturating_add(1)));
                }
            },
            Value::Null | Value::Bool(_) | Value::Number(_) | Value::String(_) => {},
        }
    }
    Ok(())
}

fn preflight_forward_json(patch: &Patch<ForwardOnly>) -> Result<(), PatchError> {
    let bytes = object_json_len(&[
        ("blobs", blob_list_json_len(&patch.forward_blobs)?),
        ("format", json_string_len(&patch.format)?),
        (
            "operations",
            operations_json_len(&patch.forward, usize::MAX)?,
        ),
        ("version", number_json_len(patch.version)),
    ])?;
    if bytes > patch.limits.json_bytes {
        return Err(PatchError::JsonLimit {
            kind: JsonLimitKind::OutputBytes,
            observed: bytes,
            limit: patch.limits.json_bytes,
        });
    }
    Ok(())
}

fn preflight_reversible_json(patch: &Patch<Reversible>) -> Result<(), PatchError> {
    let operation_bytes =
        reversible_operations_json_len(&patch.forward, &patch.reverse, usize::MAX)?;
    let bytes = object_json_len(&[
        ("format", json_string_len(&patch.format)?),
        ("forward_blobs", blob_list_json_len(&patch.forward_blobs)?),
        ("operations", operation_bytes),
        ("reverse_blobs", blob_list_json_len(&patch.reverse_blobs)?),
        ("version", number_json_len(patch.version)),
    ])?;
    if bytes > patch.limits.json_bytes {
        return Err(PatchError::JsonLimit {
            kind: JsonLimitKind::OutputBytes,
            observed: bytes,
            limit: patch.limits.json_bytes,
        });
    }
    Ok(())
}

fn operations_json_len(operations: &[PatchOperation], maximum: usize) -> Result<usize, PatchError> {
    list_json_len(
        operations
            .iter()
            .map(|operation| operation_json_len(operation, maximum)),
        maximum,
    )
}

fn reversible_operations_json_len(
    forward: &[PatchOperation],
    reverse: &[PatchOperation],
    maximum: usize,
) -> Result<usize, PatchError> {
    if forward.len() != reverse.len() {
        return Err(PatchError::OperationPairing);
    }
    list_json_len(
        forward
            .iter()
            .zip(reverse)
            .map(|(forward_operation, inverse)| {
                object_json_len(&[
                    ("forward", operation_json_len(forward_operation, maximum)?),
                    ("inverse", operation_json_len(inverse, maximum)?),
                ])
            }),
        maximum,
    )
}

fn blob_list_json_len(bundle: &BlobBundle) -> Result<usize, PatchError> {
    list_json_len(
        bundle.entries.iter().map(|(id, bytes)| {
            object_json_len(&[
                ("bytes", base64_json_string_len(bytes.len())?),
                ("sha256", json_string_len(&id.as_hex())?),
            ])
        }),
        usize::MAX,
    )
}

fn operation_json_len(operation: &PatchOperation, maximum: usize) -> Result<usize, PatchError> {
    let mut fields = vec![
        ("op", json_string_len(&operation.op)?),
        ("target", json_string_len(&operation.target)?),
        ("value", value_json_len(&operation.value, maximum)?),
    ];
    if !operation.preconditions.is_empty() {
        fields.push((
            "preconditions",
            preconditions_json_len(&operation.preconditions, maximum)?,
        ));
    }
    object_json_len(&fields)
}

fn preconditions_json_len(
    preconditions: &BTreeMap<String, Value>,
    maximum: usize,
) -> Result<usize, PatchError> {
    let mut fields = Vec::new();
    for (key, value) in preconditions {
        fields.push((key.as_str(), value_json_len(value, maximum)?));
    }
    object_json_len(&fields)
}

fn value_json_len(value: &Value, maximum: usize) -> Result<usize, PatchError> {
    let length = match value {
        Value::Null => 4,
        Value::Bool(boolean) => {
            if *boolean {
                4
            } else {
                5
            }
        },
        Value::Number(number) => number.to_string().len(),
        Value::String(text) => json_string_len(text)?,
        Value::Array(values) => list_json_len(
            values.iter().map(|child| value_json_len(child, maximum)),
            maximum,
        )?,
        Value::Object(values) => {
            let mut fields = Vec::new();
            for (key, child) in values {
                fields.push((key.as_str(), value_json_len(child, maximum)?));
            }
            object_json_len(&fields)?
        },
    };
    if length > maximum {
        return Err(PatchError::JsonLimit {
            kind: JsonLimitKind::PayloadBytes,
            observed: length,
            limit: maximum,
        });
    }
    Ok(length)
}

fn list_json_len(
    values: impl IntoIterator<Item = Result<usize, PatchError>>,
    maximum: usize,
) -> Result<usize, PatchError> {
    let mut length = 2_usize;
    let mut first = true;
    for result in values {
        let value = result?;
        if !first {
            length = checked_json_len(length, 1, maximum)?;
        }
        length = checked_json_len(length, value, maximum)?;
        first = false;
    }
    Ok(length)
}

fn object_json_len(fields: &[(&str, usize)]) -> Result<usize, PatchError> {
    let mut length = 2_usize;
    for (index, (key, value)) in fields.iter().enumerate() {
        if index != 0 {
            length = checked_json_len(length, 1, usize::MAX)?;
        }
        length = checked_json_len(length, json_string_len(key)?, usize::MAX)?;
        length = checked_json_len(length, 1, usize::MAX)?;
        length = checked_json_len(length, *value, usize::MAX)?;
    }
    Ok(length)
}

fn checked_json_len(length: usize, additional: usize, maximum: usize) -> Result<usize, PatchError> {
    let result = length
        .checked_add(additional)
        .ok_or(PatchError::JsonLimit {
            kind: JsonLimitKind::OutputBytes,
            observed: usize::MAX,
            limit: maximum,
        })?;
    if result > maximum {
        return Err(PatchError::JsonLimit {
            kind: JsonLimitKind::OutputBytes,
            observed: result,
            limit: maximum,
        });
    }
    Ok(result)
}

fn json_string_len(text: &str) -> Result<usize, PatchError> {
    let mut length = 2_usize;
    for byte in text.bytes() {
        let encoded = match byte {
            b'"' | b'\\' | b'\x08' | b'\x0c' | b'\n' | b'\r' | b'\t' => 2,
            0..=0x1f => 6,
            _ => 1,
        };
        length = checked_json_len(length, encoded, usize::MAX)?;
    }
    Ok(length)
}

fn base64_json_string_len(decoded: usize) -> Result<usize, PatchError> {
    checked_json_len(2, base64_encoded_len(decoded)?, usize::MAX)
}

const fn number_json_len(value: u16) -> usize {
    if value >= 10_000 {
        5
    } else if value >= 1_000 {
        4
    } else if value >= 100 {
        3
    } else if value >= 10 {
        2
    } else {
        1
    }
}

fn parse_canonical<T>(bytes: &[u8], limits: PatchLimits) -> Result<T, PatchError>
where
    T: for<'de> Deserialize<'de> + Serialize,
{
    if bytes.len() > limits.json_bytes {
        return Err(PatchError::JsonLimit {
            kind: JsonLimitKind::InputBytes,
            observed: bytes.len(),
            limit: limits.json_bytes,
        });
    }
    let json_value = serde_json::from_slice(bytes).map_err(PatchError::Json)?;
    validate_json_limits(&json_value, limits)?;
    let canonical = canonical_json(&json_value, limits.json_bytes)?;
    if canonical != bytes {
        return Err(PatchError::NonCanonicalJson);
    }
    serde_json::from_value(json_value).map_err(PatchError::Json)
}

fn validate_json_limits(value: &Value, limits: PatchLimits) -> Result<(), PatchError> {
    let Some(root) = value.as_object() else {
        return Ok(());
    };
    let Some(operations) = root.get("operations").and_then(Value::as_array) else {
        return Ok(());
    };
    if operations.len() > limits.operations {
        return Err(PatchError::JsonLimit {
            kind: JsonLimitKind::Operations,
            observed: operations.len(),
            limit: limits.operations,
        });
    }
    Ok(())
}

fn canonical_json<T: Serialize>(source: &T, maximum: usize) -> Result<Vec<u8>, PatchError> {
    let json_value = serde_json::to_value(source).map_err(PatchError::Json)?;
    let mut output = BoundedJsonOutput::new(maximum);
    write_json(&json_value, &mut output)?;
    Ok(output.into_inner())
}

fn write_json(value: &Value, output: &mut BoundedJsonOutput) -> Result<(), PatchError> {
    match value {
        Value::Null => output.write(b"null")?,
        Value::Bool(boolean) => {
            if *boolean {
                output.write(b"true")?;
            } else {
                output.write(b"false")?;
            }
        },
        Value::Number(number) => output.write(number.to_string().as_bytes())?,
        Value::String(text) => {
            let encoded = serde_json::to_string(text).map_err(PatchError::Json)?;
            output.write(encoded.as_bytes())?;
        },
        Value::Array(values) => {
            output.write(b"[")?;
            for (index, child) in values.iter().enumerate() {
                if index != 0 {
                    output.write(b",")?;
                }
                write_json(child, output)?;
            }
            output.write(b"]")?;
        },
        Value::Object(values) => {
            output.write(b"{")?;
            for (index, (key, child)) in values.iter().enumerate() {
                if index != 0 {
                    output.write(b",")?;
                }
                let encoded = serde_json::to_string(key).map_err(PatchError::Json)?;
                output.write(encoded.as_bytes())?;
                output.write(b":")?;
                write_json(child, output)?;
            }
            output.write(b"}")?;
        },
    }
    Ok(())
}

fn validate_version(version: u16) -> Result<(), PatchError> {
    if version == PATCH_WIRE_VERSION {
        Ok(())
    } else {
        Err(PatchError::UnsupportedVersion(version))
    }
}

fn validate_text(
    field: &'static str,
    text: &str,
    maximum: usize,
    allow_empty: bool,
) -> Result<(), PatchError> {
    if (!allow_empty && text.is_empty())
        || text.len() > maximum
        || text.chars().any(char::is_control)
    {
        return Err(PatchError::InvalidText { field });
    }
    Ok(())
}

const fn hex_nibble(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions unwrap expected success"
    )]

    use serde_json::json;

    use super::*;

    fn limits() -> BlobLimits {
        BlobLimits::new(4, 32, 64)
    }

    fn patch_limits() -> PatchLimits {
        PatchLimits::new(limits(), 4096, 8, 16, 128, 2048)
    }

    fn operation(op: &str, target: &str, value: Value) -> PatchOperation {
        PatchOperation::new(patch_limits(), op, target, BTreeMap::new(), value)
            .expect("valid operation")
    }

    #[test]
    fn reversible_wire_is_canonical_and_round_trips() {
        let mut forward = BlobBundle::new(limits());
        let forward_id = forward.insert(b"forward bytes").expect("blob fits");
        let mut reverse = BlobBundle::new(limits());
        let reverse_id = reverse.insert(b"reverse bytes").expect("blob fits");
        let patch = Patch::<Reversible>::new(
            patch_limits(),
            "org.litchi.xlsx",
            [ReversibleOperation::new(
                operation(
                    "cell.set",
                    "sheet:Summary!A1",
                    json!({"z": 2, "a": 1, "blob": forward_id.as_hex()}),
                ),
                operation(
                    "cell.set",
                    "sheet:Summary!A1",
                    json!({"blob": reverse_id.as_hex(), "a": 0}),
                ),
            )],
            forward,
            reverse,
        )
        .expect("valid patch");

        let first = patch.to_deterministic_json().expect("serialize patch");
        let decoded = Patch::<Reversible>::from_deterministic_json(&first, patch_limits())
            .expect("canonical patch parses");
        let second = decoded
            .to_deterministic_json()
            .expect("serialize decoded patch");
        assert_eq!(first, second);
        assert_eq!(
            decoded.operations()[0].value,
            json!({"a": 1, "blob": forward_id.as_hex(), "z": 2})
        );
    }

    #[test]
    fn noncanonical_json_is_rejected_even_when_it_parses() {
        let patch = Patch::<ForwardOnly>::new(
            patch_limits(),
            "org.litchi.odf",
            [operation(
                "shape.text.set",
                "page:0/shape:1",
                json!("hello"),
            )],
            BlobBundle::new(limits()),
        )
        .expect("valid patch");
        let canonical = String::from_utf8(patch.to_deterministic_json().expect("serialize patch"))
            .expect("json utf8");
        let noncanonical = canonical.replacen('{', "{ ", 1);
        assert!(matches!(
            Patch::<ForwardOnly>::from_deterministic_json(noncanonical.as_bytes(), patch_limits()),
            Err(PatchError::NonCanonicalJson)
        ));
    }

    #[test]
    fn parsing_enforces_every_explicit_json_bound() {
        let patch = Patch::<ForwardOnly>::new(
            patch_limits(),
            "org.litchi.odf",
            [operation(
                "shape.text.set",
                "page:0/shape:1",
                json!({"text": "bounded"}),
            )],
            BlobBundle::new(limits()),
        )
        .expect("valid patch");
        let wire = patch.to_deterministic_json().expect("serialize patch");
        let test_cases = [
            (
                PatchLimits::new(limits(), 1, 8, 16, 128, 2048),
                JsonLimitKind::InputBytes,
            ),
            (
                PatchLimits::new(limits(), 4096, 0, 16, 128, 2048),
                JsonLimitKind::Operations,
            ),
            (
                PatchLimits::new(limits(), 4096, 8, 1, 128, 2048),
                JsonLimitKind::ValueDepth,
            ),
            (
                PatchLimits::new(limits(), 4096, 8, 16, 2, 2048),
                JsonLimitKind::ValueStringBytes,
            ),
            (
                PatchLimits::new(limits(), 4096, 8, 16, 128, 1),
                JsonLimitKind::PayloadBytes,
            ),
        ];
        for (limits, kind) in test_cases {
            assert!(matches!(
                Patch::<ForwardOnly>::from_deterministic_json(&wire, limits),
                Err(PatchError::JsonLimit { kind: actual, .. }) if actual == kind
            ));
        }
    }

    #[test]
    fn construction_retains_and_enforces_patch_limits() {
        let permissive = PatchLimits::new(limits(), 4096, 8, 16, 128, 2048);
        let strict_strings = PatchLimits::new(limits(), 4096, 8, 16, 3, 2048);
        assert!(matches!(
            PatchOperation::new(
                strict_strings,
                "cell.set",
                "sheet:Summary!A1",
                BTreeMap::new(),
                json!("four"),
            ),
            Err(PatchError::JsonLimit {
                kind: JsonLimitKind::ValueStringBytes,
                ..
            })
        ));

        let staged = operation("cell.set", "sheet:Summary!A1", json!("value"));
        assert!(matches!(
            Patch::<ForwardOnly>::new(
                PatchLimits::new(limits(), 4096, 0, 16, 128, 2048),
                "org.litchi.xlsx",
                [staged.clone()],
                BlobBundle::new(limits()),
            ),
            Err(PatchError::JsonLimit {
                kind: JsonLimitKind::Operations,
                ..
            })
        ));
        assert!(matches!(
            Patch::<ForwardOnly>::new(
                PatchLimits::new(limits(), 4096, 8, 16, 128, 1),
                "org.litchi.xlsx",
                [staged],
                BlobBundle::new(limits()),
            ),
            Err(PatchError::JsonLimit {
                kind: JsonLimitKind::PayloadBytes,
                ..
            })
        ));
        assert!(matches!(
            Patch::<ForwardOnly>::new(
                permissive,
                "org.litchi.xlsx",
                [],
                BlobBundle::new(BlobLimits::new(1, 1, 1)),
            ),
            Err(PatchError::IncompatibleBlobLimits { .. })
        ));
    }

    #[test]
    fn semantic_object_keys_obey_value_string_bounds() {
        let strict = PatchLimits::new(limits(), 4096, 8, 16, 3, 2048);
        assert!(matches!(
            PatchOperation::new(
                strict,
                "set",
                "target",
                BTreeMap::new(),
                json!({"four": "ok"}),
            ),
            Err(PatchError::JsonLimit {
                kind: JsonLimitKind::ValueStringBytes,
                observed: 4,
                limit: 3,
            })
        ));
    }

    #[test]
    fn tight_semantic_limits_round_trip_for_both_patch_modes() {
        let tight = PatchLimits::new(BlobLimits::new(0, 0, 0), 4096, 1, 2, 1, 128);
        let forward = PatchOperation::new(
            tight,
            "operation",
            "target",
            BTreeMap::new(),
            json!({"k": "v"}),
        )
        .expect("semantic values meet tight limits");
        let forward_patch = Patch::<ForwardOnly>::new(
            tight,
            "format",
            [forward.clone()],
            BlobBundle::new(tight.blobs()),
        )
        .expect("forward patch meets tight limits");
        let forward_wire = forward_patch
            .to_deterministic_json()
            .expect("forward patch serializes");
        Patch::<ForwardOnly>::from_deterministic_json(&forward_wire, tight)
            .expect("forward patch parses under its retained limits");

        let reverse = PatchOperation::new(
            tight,
            "reverse",
            "target",
            BTreeMap::new(),
            json!({"k": "v"}),
        )
        .expect("inverse semantic values meet tight limits");
        let reversible_patch = Patch::<Reversible>::new(
            tight,
            "format",
            [ReversibleOperation::new(forward, reverse)],
            BlobBundle::new(tight.blobs()),
            BlobBundle::new(tight.blobs()),
        )
        .expect("reversible patch meets tight limits");
        let reversible_wire = reversible_patch
            .to_deterministic_json()
            .expect("reversible patch serializes");
        Patch::<Reversible>::from_deterministic_json(&reversible_wire, tight)
            .expect("reversible patch parses under its retained limits");
    }

    #[test]
    fn short_base64_preflight_reports_a_conservative_upper_bound() {
        assert_eq!(base64_decoded_upper_bound(0).expect("zero is valid"), 0);
        assert_eq!(base64_decoded_upper_bound(1).expect("one is bounded"), 3);
        assert_eq!(base64_decoded_upper_bound(2).expect("two are bounded"), 3);
        assert_eq!(base64_decoded_upper_bound(3).expect("three are bounded"), 3);
        assert_eq!(base64_decoded_upper_bound(4).expect("four are bounded"), 3);
    }

    #[test]
    fn serialization_preflight_refuses_an_oversized_envelope() {
        let serialization_bound = PatchLimits::new(limits(), 32, 8, 16, 128, 2048);
        let operation = PatchOperation::new(
            serialization_bound,
            "shape.text.set",
            "page:0/shape:1",
            BTreeMap::new(),
            json!("bounded at construction but too large for transport"),
        )
        .expect("semantic payload fits its independent payload bound");
        let patch = Patch::<ForwardOnly>::new(
            serialization_bound,
            "org.litchi.odf",
            [operation],
            BlobBundle::new(limits()),
        )
        .expect("patch construction does not allocate wire JSON");
        assert!(matches!(
            patch.to_deterministic_json(),
            Err(PatchError::JsonLimit {
                kind: JsonLimitKind::OutputBytes,
                ..
            })
        ));
    }

    #[test]
    fn blob_wire_size_is_rejected_before_identifier_or_base64_processing() {
        let bytes = "A".repeat(8);
        let invalid_id = "x".repeat(64);
        let wire = format!(
            r#"{{"blobs":[{{"bytes":"{bytes}","sha256":"{invalid_id}"}}],"format":"org.litchi.odf","operations":[],"version":1}}"#
        );
        let limits = PatchLimits::new(BlobLimits::new(1, 3, 3), 4096, 8, 16, 128, 2048);
        assert!(matches!(
            Patch::<ForwardOnly>::from_deterministic_json(wire.as_bytes(), limits),
            Err(PatchError::BlobLimit {
                kind: BlobLimitKind::BlobBytes,
                ..
            })
        ));
    }

    #[test]
    fn inverse_reverses_operation_order_and_seal_drops_reverse_data() {
        let mut forward = BlobBundle::new(limits());
        forward.insert(b"forward").expect("blob fits");
        let mut reverse = BlobBundle::new(limits());
        reverse.insert(b"private before image").expect("blob fits");
        let patch = Patch::<Reversible>::new(
            patch_limits(),
            "org.litchi.xlsx",
            [
                ReversibleOperation::new(
                    operation("first", "one", json!(1)),
                    operation("undo-first", "one", json!(0)),
                ),
                ReversibleOperation::new(
                    operation("second", "two", json!(2)),
                    operation("undo-second", "two", json!(1)),
                ),
            ],
            forward,
            reverse,
        )
        .expect("valid patch");
        let inverse = patch.inverse();
        assert_eq!(inverse.operations()[0].op, "undo-second");
        assert_eq!(inverse.operations()[1].op, "undo-first");

        let sealed = patch.seal();
        let wire = String::from_utf8(sealed.to_deterministic_json().expect("serialize sealed"))
            .expect("json utf8");
        assert!(!wire.contains("private before image"));
        assert!(!wire.contains("reverse_blobs"));
        assert_eq!(sealed.blobs().len(), 1);
    }

    #[test]
    fn blobs_are_addressed_deduplicated_and_bounded() {
        let mut bundle = BlobBundle::new(BlobLimits::new(1, 3, 3));
        let first = bundle.insert(b"abc").expect("first blob fits");
        assert_eq!(first.as_hex().len(), 64);
        assert_eq!(
            bundle.insert(b"abc").expect("duplicate deduplicates"),
            first
        );
        assert!(matches!(
            bundle.insert(b"abcd"),
            Err(PatchError::BlobLimit {
                kind: BlobLimitKind::BlobBytes,
                ..
            })
        ));
        assert_eq!(bundle.get(&first), Some(&b"abc"[..]));
    }

    #[test]
    fn history_bounds_transitions_and_preserves_undo_redo_order() {
        let mut history = History::new("base", HistoryLimits::new(2, 5));
        assert!(history.record("one", 2).expect("record first").is_empty());
        assert!(history.record("two", 2).expect("record second").is_empty());
        assert_eq!(
            history.record("three", 2).expect("evict oldest"),
            vec!["base"]
        );
        assert_eq!(history.current(), &"three");
        assert!(history.undo());
        assert_eq!(history.current(), &"two");
        assert!(history.undo());
        assert_eq!(history.current(), &"one");
        assert!(!history.undo());
        assert!(history.redo());
        assert_eq!(history.current(), &"two");
        assert_eq!(
            history.record("branch", 2).expect("branch records"),
            vec!["three"]
        );
        assert!(!history.can_redo());
        assert!(matches!(
            history.record("oversized", 6),
            Err(PatchError::HistoryWeight {
                observed: 6,
                limit: 5
            })
        ));
        assert_eq!(history.current(), &"branch");
    }

    #[test]
    fn conflict_set_retains_format_owned_deterministic_order() {
        let conflicts = ConflictSet::new(vec!["sheet:Summary!A1", "sheet:Summary!B4"]);
        assert_eq!(
            conflicts.conflicts(),
            ["sheet:Summary!A1", "sheet:Summary!B4"]
        );
        assert_eq!(conflicts.len(), 2);
        assert!(!conflicts.is_empty());
    }

    fn composition_limits() -> CompositionLimits {
        CompositionLimits::new(8, 8, 24, 8)
    }

    fn sub_edit(lineage: u64, id: &str, reads: &[&str], writes: &[&str]) -> SubEdit<u64, String> {
        SubEdit::new(
            lineage,
            composition_limits(),
            id,
            reads.iter().map(|value| (*value).to_owned()),
            writes.iter().map(|value| (*value).to_owned()),
            format!("payload:{id}"),
        )
        .expect("valid sub-edit")
    }

    #[test]
    fn patch_fingerprint_is_deterministic_and_content_sensitive() {
        let first = Patch::<ForwardOnly>::new(
            patch_limits(),
            "org.litchi.odf",
            [operation("shape.text.set", "page:0/shape:1", json!("one"))],
            BlobBundle::new(limits()),
        )
        .expect("valid first patch");
        let equivalent = Patch::<ForwardOnly>::from_deterministic_json(
            &first.to_deterministic_json().expect("serialize first"),
            patch_limits(),
        )
        .expect("parse equivalent");
        let changed = Patch::<ForwardOnly>::new(
            patch_limits(),
            "org.litchi.odf",
            [operation("shape.text.set", "page:0/shape:1", json!("two"))],
            BlobBundle::new(limits()),
        )
        .expect("valid changed patch");

        assert_eq!(
            first.fingerprint().expect("fingerprint first"),
            equivalent.fingerprint().expect("fingerprint equivalent")
        );
        assert_ne!(
            first.fingerprint().expect("fingerprint first"),
            changed.fingerprint().expect("fingerprint changed")
        );
        assert_eq!(
            first.fingerprint().expect("fingerprint first").as_hex(),
            String::from("47226755188d6f3a473b211c6d193cbccb51c07ada2dbdccd488611c719a5108")
        );
    }

    #[test]
    fn sub_edit_effects_are_canonical_and_bounded() {
        let edit = SubEdit::new(
            7_u64,
            composition_limits(),
            "canonical",
            ["shared".to_owned(), "read".to_owned()],
            ["shared".to_owned(), "write".to_owned()],
            (),
        )
        .expect("valid effects");
        assert_eq!(edit.effects().reads().collect::<Vec<_>>(), ["read"]);
        assert_eq!(
            edit.effects().writes().collect::<Vec<_>>(),
            ["shared", "write"]
        );
        assert!(matches!(
            SubEdit::new(
                7_u64,
                CompositionLimits::new(1, 1, 1, 1),
                "too-many",
                ["one".to_owned(), "two".to_owned()],
                [],
                (),
            ),
            Err(CompositionError::Limit {
                kind: CompositionLimitKind::EffectsPerSubEdit,
                observed: 2,
                limit: 1,
            })
        ));
        assert!(matches!(
            SubEdit::new(
                7_u64,
                composition_limits(),
                "invalid",
                ["bad\nkey".to_owned()],
                [],
                (),
            ),
            Err(CompositionError::InvalidEffect)
        ));
    }

    #[test]
    fn joined_sub_edits_accept_only_disjoint_effects_in_id_order() {
        let mut joined = JoinedSubEdits::new(11_u64, composition_limits());
        joined
            .join(sub_edit(11, "z-last", &["sheet:0/A1:value"], &[]))
            .expect("read-only edit joins")
            .join(sub_edit(11, "a-first", &[], &["sheet:0/B1:value"]))
            .expect("disjoint write joins");
        assert_eq!(
            joined.sub_edits().map(SubEdit::id).collect::<Vec<_>>(),
            ["a-first", "z-last"]
        );
        assert_eq!(joined.total_effects(), 2);
    }

    #[test]
    fn join_reports_read_write_conflict_and_returns_rejected_work() {
        let mut joined = JoinedSubEdits::new(13_u64, composition_limits());
        joined
            .join(sub_edit(13, "reader", &["sheet:0/A1:value"], &[]))
            .expect("first edit joins");
        let error = joined
            .join(sub_edit(13, "writer", &[], &["sheet:0/A1:value"]))
            .expect_err("read/write overlap must fail");
        assert_eq!(joined.len(), 1);
        assert_eq!(error.rejected().id(), "writer");
        let SubEditJoinFailure::Overlap(conflicts) = error.failure() else {
            panic!("expected structured overlap");
        };
        let SubEditConflict::Effect(conflict) = &conflicts.conflicts()[0] else {
            panic!("expected effect conflict");
        };
        assert_eq!(conflict.effect(), "sheet:0/A1:value");
        assert_eq!(conflict.left_access(), EffectAccess::Read);
        assert_eq!(conflict.right_access(), EffectAccess::Write);
    }

    #[test]
    fn join_refuses_to_truncate_conflict_details() {
        let strict = CompositionLimits::new(4, 4, 8, 1);
        let mut joined = JoinedSubEdits::new(17_u64, strict);
        joined
            .join(
                SubEdit::new(
                    17_u64,
                    strict,
                    "left",
                    [],
                    ["a".to_owned(), "b".to_owned()],
                    (),
                )
                .expect("left edit"),
            )
            .expect("first edit joins");
        let error = joined
            .join(
                SubEdit::new(
                    17_u64,
                    strict,
                    "right",
                    ["a".to_owned(), "b".to_owned()],
                    [],
                    (),
                )
                .expect("right edit"),
            )
            .expect_err("conflict report cannot be truncated");
        assert!(matches!(
            error.failure(),
            SubEditJoinFailure::Limit(CompositionError::Limit {
                kind: CompositionLimitKind::Conflicts,
                observed: 2,
                limit: 1,
            })
        ));
        assert_eq!(joined.len(), 1);
    }

    #[test]
    fn three_way_plan_preserves_disjoint_work_and_requires_resolution() {
        let mut unresolved_left = JoinedSubEdits::new(19_u64, composition_limits());
        unresolved_left
            .join(sub_edit(19, "left-auto", &[], &["sheet:0/A1:value"]))
            .expect("left automatic")
            .join(sub_edit(19, "left-conflict", &[], &["sheet:0/C1:value"]))
            .expect("left conflict");
        let mut unresolved_right = JoinedSubEdits::new(19_u64, composition_limits());
        unresolved_right
            .join(sub_edit(19, "right-auto", &[], &["sheet:0/B1:value"]))
            .expect("right automatic")
            .join(sub_edit(19, "right-conflict", &["sheet:0/C1:value"], &[]))
            .expect("right conflict");

        let unresolved_plan = ThreeWayMergePlan::new(unresolved_left, unresolved_right)
            .expect("plan unresolved merge");
        assert_eq!(unresolved_plan.automatic().len(), 2);
        assert_eq!(unresolved_plan.left_conflicts().len(), 1);
        assert_eq!(unresolved_plan.right_conflicts().len(), 1);
        assert!(unresolved_plan.finish().is_err());

        let mut resolved_left = JoinedSubEdits::new(19_u64, composition_limits());
        resolved_left
            .join(sub_edit(19, "left-auto", &[], &["sheet:0/A1:value"]))
            .expect("left automatic")
            .join(sub_edit(19, "left-conflict", &[], &["sheet:0/C1:value"]))
            .expect("left conflict");
        let mut resolved_right = JoinedSubEdits::new(19_u64, composition_limits());
        resolved_right
            .join(sub_edit(19, "right-auto", &[], &["sheet:0/B1:value"]))
            .expect("right automatic")
            .join(sub_edit(19, "right-conflict", &["sheet:0/C1:value"], &[]))
            .expect("right conflict");
        let mut resolved_plan =
            ThreeWayMergePlan::new(resolved_left, resolved_right).expect("plan resolved merge");
        resolved_plan.resolve(MergeChoice::Left);
        let merged = resolved_plan.finish().expect("resolved plan finishes");
        assert_eq!(
            merged.sub_edits().map(SubEdit::id).collect::<Vec<_>>(),
            ["left-auto", "left-conflict", "right-auto"]
        );
    }

    #[test]
    fn three_way_lineage_failure_returns_both_branches() {
        let left = JoinedSubEdits::<u64, ()>::new(23, composition_limits());
        let right = JoinedSubEdits::<u64, ()>::new(24, composition_limits());
        let error = ThreeWayMergePlan::new(left, right).expect_err("lineage mismatch");
        assert_eq!(error.failure(), &ThreeWayMergeFailure::DifferentLineage);
        let (recovered_left, recovered_right) = error.into_branches();
        assert_eq!(recovered_left.lineage(), &23);
        assert_eq!(recovered_right.lineage(), &24);
    }
}
