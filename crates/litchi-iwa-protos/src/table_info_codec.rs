//! Strict private Buffa projection for `TST.TableInfoArchive` model ownership.
//!
//! The strict raw-wire pass canonicalizes every visited field framing before
//! Buffa observes the source. It selects only the required model reference and
//! its non-zero identifier; strict preflight additionally projects the
//! presence-preserving drawable lock state from the required `super` envelope.
//! All other table metadata remains caller-owned bytes and is never
//! materialized or re-encoded.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict preflight intentionally precedes the low-level wire reader it consumes."
)]

use std::{fmt, num::NonZeroU64};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_table_info_generated::LitchiIwaProjection as projection;

const TABLE_SUPER_FIELD: u32 = 1;
const TABLE_MODEL_FIELD: u32 = 2;
const DRAWABLE_LOCKED_FIELD: u32 = 5;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const MAX_RECURSION_LIMIT: u32 = 64;

/// Explicit finite resource policy for one `TableInfo` model-reference decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_output_bytes: usize,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
}

impl DecodeOptions {
    /// Build a finite bytes/fields/work/nesting policy.
    ///
    /// Work accounts for both the strict scan and Buffa's deferred model
    /// reference access. The selected nested reference is charged separately
    /// because it is scanned and forced after the outer message.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_output_bytes: max_message_bytes,
            max_allocations: usize::MAX,
            max_retained_bytes: usize::MAX,
            max_scratch_bytes: usize::MAX,
        }
    }

    /// Build conservative finite limits from one known source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self {
            max_message_bytes: bytes,
            max_fields: bytes.checked_mul(8).unwrap_or(usize::MAX).max(1),
            max_work_bytes: bytes.checked_mul(32).unwrap_or(usize::MAX).max(1),
            recursion_limit: 8,
            max_output_bytes: bytes.checked_mul(2).unwrap_or(usize::MAX).max(1),
            max_allocations: usize::MAX,
            max_retained_bytes: usize::MAX,
            max_scratch_bytes: bytes.checked_mul(8).unwrap_or(usize::MAX).max(1),
        }
    }

    /// Replace the candidate-output ceiling used by prepared rewrites.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Return the candidate-output ceiling used by prepared rewrites.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Replace the logical allocation ceiling used by prepared rewrites.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }

    /// Return the logical allocation ceiling used by prepared rewrites.
    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }

    /// Replace the retained-candidate ceiling used by prepared rewrites.
    #[must_use]
    pub const fn with_max_retained_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_bytes = maximum;
        self
    }

    /// Return the retained-candidate ceiling used by prepared rewrites.
    #[must_use]
    pub const fn max_retained_bytes(self) -> usize {
        self.max_retained_bytes
    }

    /// Replace the scratch-byte ceiling used by prepared rewrites.
    #[must_use]
    pub const fn with_max_scratch_bytes(mut self, maximum: usize) -> Self {
        self.max_scratch_bytes = maximum;
        self
    }

    /// Return the scratch-byte ceiling used by prepared rewrites.
    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }

    /// Replace the input-message ceiling.
    #[must_use]
    pub const fn with_max_message_bytes(mut self, maximum: usize) -> Self {
        self.max_message_bytes = maximum;
        self
    }

    /// Return the input-message ceiling.
    #[must_use]
    pub const fn max_message_bytes(self) -> usize {
        self.max_message_bytes
    }

    /// Replace the strict field ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Return the strict field ceiling.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Replace the strict work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Return the strict work ceiling.
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    /// Replace the protobuf recursion ceiling.
    #[must_use]
    pub const fn with_recursion_limit(mut self, maximum: u32) -> Self {
        self.recursion_limit = maximum;
        self
    }

    /// Return the protobuf recursion ceiling.
    #[must_use]
    pub const fn recursion_limit(self) -> u32 {
        self.recursion_limit
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }

    fn descend(self) -> Result<Self, DecodeError> {
        let recursion_limit = self
            .recursion_limit
            .checked_sub(1)
            .ok_or_else(DecodeError::recursion_limit)?;
        Ok(Self {
            recursion_limit,
            ..self
        })
    }
}

/// Typed non-zero reference to a native `TST.TableModelArchive` object.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableModelReference {
    identifier: NonZeroU64,
}

/// Generated-free facts from one `TST.TableInfoArchive`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableInfoSnapshot {
    table_model: TableModelReference,
    locked: Option<bool>,
}

impl TableInfoSnapshot {
    /// Required non-zero native `TST.TableModelArchive` object reference.
    #[must_use]
    pub const fn table_model(self) -> TableModelReference {
        self.table_model
    }

    /// Explicit `TSD.DrawableArchive.locked` state, preserving source presence.
    ///
    /// `None` means the native field was absent; it does not apply a UI default.
    #[must_use]
    pub const fn locked(self) -> Option<bool> {
        self.locked
    }
}

/// Presence-aware lock value for a prepared `TableInfo` rewrite.
///
/// A boolean request represents the semantic lock state.  When the requested
/// state is unlocked, an absent source field remains absent; an explicitly
/// present `false` field remains explicit.  [`Self::explicit`] is available to
/// callers that need to request an exact protobuf presence (`None` removes the
/// field).  This keeps the native absent/false/true distinction intact without
/// exposing the generated `TSD.DrawableArchive` type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableInfoLockWrite {
    locked: Option<bool>,
    preserve_absent_unlocked: bool,
    expected_fingerprint: Option<u64>,
}

impl TableInfoLockWrite {
    /// Build a semantic lock write from either `bool`, `Option<bool>`, or an
    /// existing request.  `bool` preserves an absent unlocked source field;
    /// `Option<bool>` requests exact protobuf presence.
    #[must_use]
    pub fn new(value: impl Into<Self>) -> Self {
        value.into()
    }

    /// Build a semantic lock write while preserving an absent unlocked field.
    #[must_use]
    pub const fn from_locked(locked: bool) -> Self {
        Self {
            locked: Some(locked),
            preserve_absent_unlocked: true,
            expected_fingerprint: None,
        }
    }

    /// Build an exact presence-preserving write.  `None` removes the lock
    /// field, while `Some(false)` retains an explicit false field.
    #[must_use]
    pub const fn explicit(locked: Option<bool>) -> Self {
        Self {
            locked,
            preserve_absent_unlocked: false,
            expected_fingerprint: None,
        }
    }

    /// Build a request with a complete-source fingerprint expectation.
    #[must_use]
    pub fn with_fingerprint(expected_fingerprint: u64, value: impl Into<Self>) -> Self {
        value.into().expecting_fingerprint(expected_fingerprint)
    }

    /// Add a complete-source fingerprint expectation to this request.
    #[must_use]
    pub const fn expecting_fingerprint(mut self, expected_fingerprint: u64) -> Self {
        self.expected_fingerprint = Some(expected_fingerprint);
        self
    }

    /// Return the requested lock value, including requested presence.
    #[must_use]
    pub const fn locked(self) -> Option<bool> {
        self.locked
    }

    /// Return the requested semantic lock state, treating absence as unlocked.
    #[must_use]
    pub const fn is_locked(self) -> bool {
        matches!(self.locked, Some(true))
    }

    /// Return whether an unlocked source may remain absent.
    #[must_use]
    pub const fn preserves_absent_unlocked(self) -> bool {
        self.preserve_absent_unlocked
    }

    /// Return the optional source fingerprint expectation.
    #[must_use]
    pub const fn expected_fingerprint(self) -> Option<u64> {
        self.expected_fingerprint
    }
}

impl From<bool> for TableInfoLockWrite {
    fn from(locked: bool) -> Self {
        Self::from_locked(locked)
    }
}

impl From<Option<bool>> for TableInfoLockWrite {
    fn from(locked: Option<bool>) -> Self {
        Self::explicit(locked)
    }
}

impl TableModelReference {
    /// Native model object identifier, proven non-zero by strict preflight.
    #[must_use]
    pub const fn identifier(self) -> NonZeroU64 {
        self.identifier
    }
}

/// Failure from `TableInfo` strict preflight or its Buffa cross-check.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

/// A content-free wire-resource classification for [`DecodeError`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    /// The input or configured message-byte ceiling could not be honored.
    Bytes {
        /// Input length when it was known at the failure point.
        observed: Option<usize>,
        /// Applied byte ceiling when known.
        maximum: Option<usize>,
    },
    /// The configured or enforced protobuf nesting ceiling was exceeded.
    Nesting {
        /// Configured nesting value when the profile itself was invalid.
        observed: Option<u32>,
        /// Applied nesting ceiling when known.
        maximum: Option<u32>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    WireResourceLimit(WireResourceLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    ZeroIdentifier(&'static str),
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
    OutputLimit { observed: usize, maximum: usize },
    AllocationLimit { observed: usize, maximum: usize },
    RetainedLimit { observed: usize, maximum: usize },
    ScratchLimit { observed: usize, maximum: usize },
    Allocation { amount: usize },
    FingerprintMismatch,
    Projection,
}

impl DecodeError {
    fn recursion_limit() -> Self {
        Self::wire_resource_limit_error(WireResourceLimit::Nesting {
            observed: None,
            maximum: None,
        })
    }

    const fn wire_resource_limit_error(limit: WireResourceLimit) -> Self {
        Self {
            kind: DecodeErrorKind::WireResourceLimit(limit),
        }
    }

    fn with_recursion_limit_context(mut self, maximum: u32) -> Self {
        if let DecodeErrorKind::WireResourceLimit(WireResourceLimit::Nesting {
            observed,
            maximum: None,
        }) = self.kind
        {
            self.kind = DecodeErrorKind::WireResourceLimit(WireResourceLimit::Nesting {
                observed,
                maximum: Some(maximum),
            });
        }
        self
    }

    const fn missing_required(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::MissingRequired(field),
        }
    }

    const fn duplicate_singular(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateSingular(field),
        }
    }

    const fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason),
        }
    }

    const fn zero_identifier(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::ZeroIdentifier(field),
        }
    }

    const fn field_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::FieldLimit { observed, maximum },
        }
    }

    const fn work_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::WorkLimit { observed, maximum },
        }
    }

    const fn output_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::OutputLimit { observed, maximum },
        }
    }

    const fn allocation_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::AllocationLimit { observed, maximum },
        }
    }

    const fn retained_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::RetainedLimit { observed, maximum },
        }
    }

    const fn scratch_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::ScratchLimit { observed, maximum },
        }
    }

    const fn allocation(amount: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation { amount },
        }
    }

    const fn fingerprint_mismatch() -> Self {
        Self {
            kind: DecodeErrorKind::FingerprintMismatch,
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    /// Required schema field absent from the source, when applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        let DecodeErrorKind::MissingRequired(field) = self.kind else {
            return None;
        };
        Some(field)
    }

    /// Singular schema field repeated in the source, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        let DecodeErrorKind::DuplicateSingular(field) = self.kind else {
            return None;
        };
        Some(field)
    }

    /// Stable canonicality failure reason, when applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        let DecodeErrorKind::NonCanonical(reason) = self.kind else {
            return None;
        };
        Some(reason)
    }

    /// Reference field carrying a forbidden zero identifier, when applicable.
    #[must_use]
    pub const fn zero_identifier_field(&self) -> Option<&'static str> {
        let DecodeErrorKind::ZeroIdentifier(field) = self.kind else {
            return None;
        };
        Some(field)
    }

    /// Observed and configured field counts for a field-limit failure.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::FieldLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Observed and configured work bytes for a work-limit failure.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::WorkLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Observed and configured candidate-output bytes for a rewrite limit.
    #[must_use]
    pub const fn output_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::OutputLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Observed and configured logical output allocations for a rewrite limit.
    #[must_use]
    pub const fn allocation_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::AllocationLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Observed and configured retained bytes for a rewrite limit.
    #[must_use]
    pub const fn retained_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::RetainedLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Observed and configured scratch bytes for a rewrite limit.
    #[must_use]
    pub const fn scratch_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::ScratchLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Requested output allocation when reservation failed.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        let DecodeErrorKind::Allocation { amount } = self.kind else {
            return None;
        };
        Some(amount)
    }

    /// Whether a prepared rewrite observed a stale source fingerprint.
    #[must_use]
    pub const fn is_fingerprint_mismatch(&self) -> bool {
        matches!(self.kind, DecodeErrorKind::FingerprintMismatch)
    }

    /// Wire byte/nesting resource failure, independent of Buffa error text.
    #[must_use]
    pub const fn wire_resource_limit(&self) -> Option<WireResourceLimit> {
        let DecodeErrorKind::WireResourceLimit(limit) = self.kind else {
            return None;
        };
        Some(limit)
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::WireResourceLimit(WireResourceLimit::Bytes { .. }) => {
                formatter.write_str("Numbers TableInfo wire byte limit exceeded")
            },
            DecodeErrorKind::WireResourceLimit(WireResourceLimit::Nesting { .. }) => {
                formatter.write_str("Numbers TableInfo wire nesting limit exceeded")
            },
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::ZeroIdentifier(field) => write!(formatter, "{field} is zero"),
            DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Numbers TableInfo projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Numbers TableInfo projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::OutputLimit { observed, maximum } => write!(
                formatter,
                "TableInfo lock rewrite produced {observed} output bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::AllocationLimit { observed, maximum } => write!(
                formatter,
                "TableInfo lock rewrite requires {observed} allocations; maximum is {maximum}"
            ),
            DecodeErrorKind::RetainedLimit { observed, maximum } => write!(
                formatter,
                "TableInfo lock rewrite retains {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::ScratchLimit { observed, maximum } => write!(
                formatter,
                "TableInfo lock rewrite requires {observed} scratch bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Allocation { amount } => write!(
                formatter,
                "cannot allocate TableInfo lock rewrite output for {amount} bytes"
            ),
            DecodeErrorKind::FingerprintMismatch => {
                formatter.write_str("TableInfo lock rewrite source fingerprint mismatch")
            },
            DecodeErrorKind::Projection => formatter.write_str(
                "Numbers TableInfo strict preflight disagrees with the Buffa projection",
            ),
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        match error {
            buffa::DecodeError::MessageTooLarge => {
                Self::wire_resource_limit_error(WireResourceLimit::Bytes {
                    observed: None,
                    maximum: None,
                })
            },
            buffa::DecodeError::RecursionLimitExceeded => Self::recursion_limit(),
            error => Self {
                kind: DecodeErrorKind::Wire(error),
            },
        }
    }
}

/// Decode model ownership and drawable lock state from one `TableInfo` payload.
///
/// Every root field is strictly scanned for canonical protobuf framing. The
/// required `super` envelope is strictly scanned for the optional lock field;
/// field 2's deferred reference is then forced once after strict preflight.
/// All unselected metadata remains opaque in the caller-owned source
/// representation.
pub fn decode_table_info(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableInfoSnapshot, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    decode_table_info_with_budget(source, options, &mut budget)
}

fn decode_table_info_with_budget(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<TableInfoSnapshot, DecodeError> {
    let strict = preflight_table_info(source, options, budget)
        .map_err(|error| error.with_recursion_limit_context(options.recursion_limit))?;

    let view: projection::TableInfoArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)
        .map_err(|error| error.with_recursion_limit_context(options.recursion_limit))?;
    let super_ = view
        .super_
        .get()
        .map_err(DecodeError::from)
        .map_err(|error| error.with_recursion_limit_context(options.recursion_limit))?
        .ok_or_else(|| DecodeError::missing_required("TST.TableInfoArchive.super"))?;
    if !view.has_table_model() {
        return Err(DecodeError::missing_required(
            "TST.TableInfoArchive.table_model",
        ));
    }
    let model = view
        .table_model
        .get()
        .map_err(DecodeError::from)
        .map_err(|error| error.with_recursion_limit_context(options.recursion_limit))?
        .ok_or_else(|| DecodeError::missing_required("TST.TableInfoArchive.table_model"))?;
    if !model.has_identifier() {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    let projected_model = TableModelReference {
        identifier: NonZeroU64::new(model.identifier)
            .ok_or_else(|| DecodeError::zero_identifier("TSP.Reference.identifier"))?,
    };
    let projected = TableInfoSnapshot {
        table_model: projected_model,
        locked: super_.locked,
    };
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

/// Decode only the required `TableInfo` table-model reference.
///
/// This compatibility convenience retains strict lock-state validation.
pub fn decode_table_model_reference(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableModelReference, DecodeError> {
    Ok(decode_table_info(source, options)?.table_model())
}

/// Presence-aware lock rewrite request accepted by the prepared codec.
pub type TableInfoLockStateWrite = TableInfoLockWrite;

/// Aggregate accounting produced by a prepared `TableInfo` lock rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    changed: bool,
    source_fingerprint: u64,
}

impl RewriteReport {
    /// Source bytes inspected during preparation.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Exact candidate bytes emitted by execution.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Aggregate strict field visits charged by the prepared operation.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate wire work charged by the prepared operation.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum strict wire depth observed by the prepared operation.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Number of logical output reservations required by execution.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Candidate bytes retained by the returned output.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Logical scratch bytes traversed or retained during the operation.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Whether the requested lock representation differs from the source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Stable FNV-1a fingerprint of the prepared source payload.
    #[must_use]
    pub const fn source_fingerprint(self) -> u64 {
        self.source_fingerprint
    }
}

/// Exact accounting ceilings required by a prepared lock rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
    /// Return ceilings that accept exactly these requirements.
    #[must_use]
    pub const fn exact(self) -> RewriteExecutionLimits {
        RewriteExecutionLimits {
            input_bytes: self.input_bytes,
            output_bytes: self.output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
        }
    }

    /// Compatibility spelling used by package transaction owners.
    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        self.exact()
    }

    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Caller-provided ceilings replayed before the output buffer is allocated.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    /// Build exact limits from prepared requirements.
    #[must_use]
    pub const fn exact(requirements: RewriteExecutionRequirements) -> Self {
        requirements.exact()
    }

    /// Start with no additional execution restriction.
    #[must_use]
    pub const fn unrestricted() -> Self {
        Self {
            input_bytes: usize::MAX,
            output_bytes: usize::MAX,
            fields: usize::MAX,
            work_bytes: usize::MAX,
            max_depth: u32::MAX,
            allocations: usize::MAX,
            retained_bytes: usize::MAX,
            scratch_bytes: usize::MAX,
        }
    }

    #[must_use]
    pub const fn with_input_bytes(mut self, maximum: usize) -> Self {
        self.input_bytes = maximum;
        self
    }

    #[must_use]
    pub const fn with_output_bytes(mut self, maximum: usize) -> Self {
        self.output_bytes = maximum;
        self
    }

    #[must_use]
    pub const fn with_fields(mut self, maximum: usize) -> Self {
        self.fields = maximum;
        self
    }

    #[must_use]
    pub const fn with_work_bytes(mut self, maximum: usize) -> Self {
        self.work_bytes = maximum;
        self
    }

    #[must_use]
    pub const fn with_max_depth(mut self, maximum: u32) -> Self {
        self.max_depth = maximum;
        self
    }

    #[must_use]
    pub const fn with_allocations(mut self, maximum: usize) -> Self {
        self.allocations = maximum;
        self
    }

    #[must_use]
    pub const fn with_retained_bytes(mut self, maximum: usize) -> Self {
        self.retained_bytes = maximum;
        self
    }

    #[must_use]
    pub const fn with_scratch_bytes(mut self, maximum: usize) -> Self {
        self.scratch_bytes = maximum;
        self
    }
}

/// Candidate bytes and the report produced by one prepared execution.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RewriteOutput {
    bytes: Vec<u8>,
    report: RewriteReport,
}

impl RewriteOutput {
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    #[must_use]
    pub fn output(&self) -> &[u8] {
        &self.bytes
    }

    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    #[must_use]
    pub fn into_output(self) -> Vec<u8> {
        self.bytes
    }

    #[must_use]
    pub const fn report(&self) -> RewriteReport {
        self.report
    }
}

/// Compatibility aliases with an explicit TableInfo lock prefix.
pub type TableInfoLockRewriteReport = RewriteReport;
pub type TableInfoLockRewriteRequirements = RewriteExecutionRequirements;
pub type TableInfoLockRewriteLimits = RewriteExecutionLimits;
pub type TableInfoLockRewriteOutput = RewriteOutput;

/// Borrowed, strictly preflighted `TSD.DrawableArchive.locked` rewrite.
#[derive(Debug, Clone, Copy)]
pub struct PreparedTableInfoLockRewrite<'source> {
    source: &'source [u8],
    write: TableInfoLockWrite,
    target_locked: Option<bool>,
    options: DecodeOptions,
    report: RewriteReport,
    requirements: RewriteExecutionRequirements,
    current: TableInfoSnapshot,
}

impl<'source> PreparedTableInfoLockRewrite<'source> {
    /// Return preparation accounting without allocating candidate output.
    #[must_use]
    pub const fn prepare_report(&self) -> RewriteReport {
        self.report
    }

    /// Return exact ceilings consumed by execution.
    #[must_use]
    pub const fn execution_requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Return the stable source fingerprint captured at preparation time.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.report.source_fingerprint
    }

    /// Execute after replaying every prepared ceiling.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_rewrite_execution_limits(self.requirements, limits)?;
        let mut output = reserve_rewrite_output(self.requirements.output_bytes)?;
        if self.report.changed {
            emit_lock_rewrite(self.source, self.write, self.options, &mut output)?;
        } else {
            append_rewrite_bytes(&mut output, self.source)?;
        }
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::projection());
        }

        if self.report.changed {
            // Candidate readback was measured during preparation. Reuse the
            // exact prepared resource ceilings instead of resetting the
            // allocation/retained/scratch axes to unbounded values.
            let readback_options = DecodeOptions::for_source(&output)
                .with_max_output_bytes(output.len())
                .with_max_allocations(self.requirements.allocations)
                .with_max_retained_bytes(self.requirements.retained_bytes)
                .with_max_scratch_bytes(self.requirements.scratch_bytes);
            validate_decode_input(&output, readback_options)?;
            let mut readback_budget = Budget::new(readback_options);
            let readback =
                decode_table_info_with_budget(&output, readback_options, &mut readback_budget)?;
            if readback.locked() != self.target_locked
                || readback.table_model() != self.current.table_model()
            {
                return Err(DecodeError::projection());
            }
        } else if self.current.locked() != self.target_locked {
            return Err(DecodeError::projection());
        }

        Ok(RewriteOutput {
            bytes: output,
            report: self.report,
        })
    }
}

/// Prepare a strict raw-preserving lock rewrite without allocating output.
pub fn prepare_table_info_lock_rewrite<'source>(
    source: &'source [u8],
    write: impl Into<TableInfoLockWrite>,
    options: DecodeOptions,
) -> Result<PreparedTableInfoLockRewrite<'source>, DecodeError> {
    let write = write.into();
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let current = decode_table_info_with_budget(source, options, &mut budget)?;
    if write
        .expected_fingerprint()
        .is_some_and(|expected| expected != table_info_source_fingerprint(source))
    {
        return Err(DecodeError::fingerprint_mismatch());
    }

    let target_locked = effective_lock_value(current.locked(), write);
    let changed = target_locked != current.locked();
    let source_fields = budget.fields;
    let output_bytes = if changed {
        let measure = measure_lock_rewrite(source, write, options, &mut budget)?;
        let output_bytes = measure.output_bytes;
        if output_bytes > options.max_output_bytes {
            return Err(DecodeError::output_limit(
                output_bytes,
                options.max_output_bytes,
            ));
        }

        // Candidate readback is charged before execution.  Its field count
        // differs only by the selected optional lock field; all unknown and
        // required source framing is retained byte-for-byte.
        let candidate_fields =
            adjusted_field_count(source_fields, current.locked(), target_locked)?;
        budget.charge_fields(candidate_fields)?;
        budget.charge_message(output_bytes)?;
        budget.charge_message(measure.drawable_bytes)?;
        budget.charge_message(measure.model_bytes)?;

        // Execution scans only the root and selected drawable before writing;
        // charge that traversal while no candidate buffer exists.
        let _ = measure_lock_rewrite(source, write, options, &mut budget)?;
        output_bytes
    } else {
        if source.len() > options.max_output_bytes {
            return Err(DecodeError::output_limit(
                source.len(),
                options.max_output_bytes,
            ));
        }
        source.len()
    };

    let scratch_bytes = if changed {
        source
            .len()
            .checked_add(output_bytes)
            .ok_or_else(DecodeError::projection)?
    } else {
        source.len()
    };
    let requirements = RewriteExecutionRequirements {
        input_bytes: source.len(),
        output_bytes,
        fields: budget.fields,
        work_bytes: budget.work_bytes,
        max_depth: budget.max_depth,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes,
    };
    check_prepare_rewrite_limits(requirements, options)?;
    let report = RewriteReport {
        input_bytes: source.len(),
        output_bytes,
        fields: requirements.fields,
        work_bytes: requirements.work_bytes,
        max_depth: requirements.max_depth,
        allocations: requirements.allocations,
        retained_bytes: requirements.retained_bytes,
        scratch_bytes: requirements.scratch_bytes,
        changed,
        source_fingerprint: table_info_source_fingerprint(source),
    };
    Ok(PreparedTableInfoLockRewrite {
        source,
        write,
        target_locked,
        options,
        report,
        requirements,
        current,
    })
}

/// Prepare a lock rewrite guarded by a complete-source fingerprint.
pub fn prepare_table_info_lock_rewrite_with_fingerprint<'source>(
    source: &'source [u8],
    expected_fingerprint: u64,
    write: impl Into<TableInfoLockWrite>,
    options: DecodeOptions,
) -> Result<PreparedTableInfoLockRewrite<'source>, DecodeError> {
    prepare_table_info_lock_rewrite(
        source,
        TableInfoLockWrite::with_fingerprint(expected_fingerprint, write),
        options,
    )
}

/// One-shot strict lock rewrite through the prepared contract.
pub fn rewrite_table_info_lock(
    source: &[u8],
    write: impl Into<TableInfoLockWrite>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let prepared = prepare_table_info_lock_rewrite(source, write, options)?;
    Ok(prepared
        .execute(prepared.execution_requirements().exact())?
        .into_bytes())
}

/// One-shot lock rewrite returning exact aggregate accounting.
pub fn rewrite_table_info_lock_with_report(
    source: &[u8],
    write: impl Into<TableInfoLockWrite>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let prepared = prepare_table_info_lock_rewrite(source, write, options)?;
    let output = prepared.execute(prepared.execution_requirements().exact())?;
    let report = output.report();
    Ok((output.into_bytes(), report))
}

/// Stable allocation-free fingerprint for a complete TableInfo payload.
#[must_use]
pub fn table_info_source_fingerprint(source: &[u8]) -> u64 {
    let mut fingerprint = 14_695_981_039_346_656_037u64;
    for byte in source {
        fingerprint ^= u64::from(*byte);
        fingerprint = fingerprint.wrapping_mul(1_099_511_628_211u64);
    }
    fingerprint
}

/// Compatibility spelling used by prepared package owners.
#[must_use]
pub fn table_info_fingerprint(source: &[u8]) -> u64 {
    table_info_source_fingerprint(source)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let max_buffa_message_bytes =
        usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_conversion| {
            DecodeError::wire_resource_limit_error(WireResourceLimit::Bytes {
                observed: None,
                maximum: None,
            })
        })?;
    if options.max_message_bytes > max_buffa_message_bytes {
        return Err(DecodeError::wire_resource_limit_error(
            WireResourceLimit::Bytes {
                observed: None,
                maximum: Some(max_buffa_message_bytes),
            },
        ));
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::wire_resource_limit_error(
            WireResourceLimit::Bytes {
                observed: Some(source.len()),
                maximum: Some(options.max_message_bytes),
            },
        ));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::wire_resource_limit_error(
            WireResourceLimit::Nesting {
                observed: Some(options.recursion_limit),
                maximum: Some(MAX_RECURSION_LIMIT),
            },
        ));
    }
    Ok(())
}

#[derive(Debug)]
struct Budget {
    fields: usize,
    work_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_depth: u32,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
            max_depth: 1,
        }
    }

    fn charge_field(&mut self) -> Result<(), DecodeError> {
        self.charge_fields(1)
    }

    fn charge_fields(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self
            .fields
            .checked_add(amount)
            .ok_or_else(|| DecodeError::field_limit(usize::MAX, self.max_fields))?;
        if observed > self.max_fields {
            return Err(DecodeError::field_limit(observed, self.max_fields));
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_message(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let strict_and_projection = bytes
            .checked_mul(2)
            .ok_or_else(|| DecodeError::work_limit(usize::MAX, self.max_work_bytes))?;
        let observed = self
            .work_bytes
            .checked_add(strict_and_projection)
            .ok_or_else(|| DecodeError::work_limit(usize::MAX, self.max_work_bytes))?;
        if observed > self.max_work_bytes {
            return Err(DecodeError::work_limit(observed, self.max_work_bytes));
        }
        self.work_bytes = observed;
        Ok(())
    }
}

fn preflight_table_info(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<TableInfoSnapshot, DecodeError> {
    budget.charge_message(source.len())?;
    let nested_options = options.descend()?;
    let mut model = None;
    let mut locked = None;
    let mut saw_super = false;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            TABLE_SUPER_FIELD => {
                if saw_super {
                    return Err(DecodeError::duplicate_singular(
                        "TST.TableInfoArchive.super",
                    ));
                }
                saw_super = true;
                budget.max_depth = budget.max_depth.max(2);
                locked = Some(preflight_drawable(
                    field.length_delimited()?,
                    nested_options,
                    budget,
                )?);
            },
            TABLE_MODEL_FIELD => {
                if model.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TST.TableInfoArchive.table_model",
                    ));
                }
                budget.max_depth = budget.max_depth.max(2);
                model = Some(preflight_reference(
                    field.length_delimited()?,
                    nested_options,
                    budget,
                )?);
            },
            _ => {},
        }
    }
    if !saw_super {
        return Err(DecodeError::missing_required("TST.TableInfoArchive.super"));
    }
    Ok(TableInfoSnapshot {
        table_model: model
            .ok_or_else(|| DecodeError::missing_required("TST.TableInfoArchive.table_model"))?,
        locked: locked.ok_or_else(DecodeError::projection)?,
    })
}

fn preflight_drawable(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Option<bool>, DecodeError> {
    budget.charge_message(source.len())?;
    let mut locked = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        if field.number != DRAWABLE_LOCKED_FIELD {
            continue;
        }
        if locked.is_some() {
            return Err(DecodeError::duplicate_singular(
                "TSD.DrawableArchive.locked",
            ));
        }
        locked = Some(require_canonical_bool(field.varint()?)?);
    }
    Ok(locked)
}

fn preflight_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<TableModelReference, DecodeError> {
    budget.charge_message(source.len())?;
    let mut identifier = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        if field.number != REFERENCE_IDENTIFIER_FIELD {
            continue;
        }
        if identifier.is_some() {
            return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
        }
        identifier = Some(
            NonZeroU64::new(field.varint()?)
                .ok_or_else(|| DecodeError::zero_identifier("TSP.Reference.identifier"))?,
        );
    }
    Ok(TableModelReference {
        identifier: identifier
            .ok_or_else(|| DecodeError::missing_required("TSP.Reference.identifier"))?,
    })
}

#[derive(Clone, Copy, Debug)]
struct FieldSpan<'source> {
    field: StrictField<'source>,
    start: usize,
    end: usize,
}

/// Visit complete strict field spans while retaining the original source.
///
/// The visitor is used only for the selected `super` envelope.  It therefore
/// never materializes unknown fields, and callers can copy their original
/// framing directly into a candidate buffer.
fn visit_field_spans<'source, F>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    mut visitor: F,
) -> Result<(), DecodeError>
where
    F: FnMut(FieldSpan<'source>) -> Result<(), DecodeError>,
{
    budget.charge_message(source.len())?;
    let mut remaining = source;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let field = next_strict_field(&mut remaining, options, budget)?
            .ok_or(buffa::DecodeError::UnexpectedEof)?;
        let end = source.len() - remaining.len();
        visitor(FieldSpan { field, start, end })?;
    }
    Ok(())
}

#[derive(Clone, Copy, Debug)]
struct RewriteMeasure {
    output_bytes: usize,
    drawable_bytes: usize,
    model_bytes: usize,
}

fn effective_lock_value(current: Option<bool>, write: TableInfoLockWrite) -> Option<bool> {
    match write.locked() {
        Some(value) if !value && write.preserves_absent_unlocked() && current.is_none() => None,
        Some(value) => Some(value),
        None => None,
    }
}

fn adjusted_field_count(
    source_fields: usize,
    current: Option<bool>,
    target: Option<bool>,
) -> Result<usize, DecodeError> {
    match (current, target) {
        (None, Some(_)) => source_fields
            .checked_add(1)
            .ok_or_else(|| DecodeError::field_limit(usize::MAX, usize::MAX)),
        (Some(_), None) => source_fields
            .checked_sub(1)
            .ok_or_else(DecodeError::projection),
        _ => Ok(source_fields),
    }
}

fn measure_lock_rewrite(
    source: &[u8],
    write: TableInfoLockWrite,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<RewriteMeasure, DecodeError> {
    let nested_options = options.descend()?;
    let mut output_bytes = 0usize;
    let mut drawable_bytes = None;
    let mut model_bytes = None;
    let mut saw_super = false;
    let mut saw_model = false;
    budget.charge_message(source.len())?;
    let mut remaining = source;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let field = next_strict_field(&mut remaining, options, budget)?
            .ok_or(buffa::DecodeError::UnexpectedEof)?;
        let end = source.len() - remaining.len();
        match field.number {
            TABLE_SUPER_FIELD => {
                if saw_super {
                    return Err(DecodeError::duplicate_singular(
                        "TST.TableInfoArchive.super",
                    ));
                }
                saw_super = true;
                let payload = field.length_delimited()?;
                let nested = measure_drawable_rewrite(payload, write, nested_options, budget)?;
                drawable_bytes = Some(nested);
                let field_length = encoded_length_delimited_field_len(TABLE_SUPER_FIELD, nested)?;
                output_bytes = output_bytes
                    .checked_add(field_length)
                    .ok_or_else(DecodeError::projection)?;
            },
            TABLE_MODEL_FIELD => {
                if saw_model {
                    return Err(DecodeError::duplicate_singular(
                        "TST.TableInfoArchive.table_model",
                    ));
                }
                saw_model = true;
                let payload = field.length_delimited()?;
                model_bytes = Some(payload.len());
                output_bytes = output_bytes
                    .checked_add(end - start)
                    .ok_or_else(DecodeError::projection)?;
            },
            _ => {
                output_bytes = output_bytes
                    .checked_add(end - start)
                    .ok_or_else(DecodeError::projection)?;
            },
        }
    }
    if !saw_super || !saw_model {
        return Err(DecodeError::projection());
    }
    Ok(RewriteMeasure {
        output_bytes,
        drawable_bytes: drawable_bytes.ok_or_else(DecodeError::projection)?,
        model_bytes: model_bytes.ok_or_else(DecodeError::projection)?,
    })
}

fn measure_drawable_rewrite(
    source: &[u8],
    write: TableInfoLockWrite,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    let mut output_bytes = 0usize;
    let mut current = None;
    visit_field_spans(source, options, budget, |span| {
        if span.field.number == DRAWABLE_LOCKED_FIELD {
            if current.is_some() {
                return Err(DecodeError::duplicate_singular(
                    "TSD.DrawableArchive.locked",
                ));
            }
            current = Some(require_canonical_bool(span.field.varint()?)?);
        } else {
            output_bytes = output_bytes
                .checked_add(span.end - span.start)
                .ok_or_else(DecodeError::projection)?;
        }
        Ok(())
    })?;
    if let Some(value) = effective_lock_value(current, write) {
        output_bytes = output_bytes
            .checked_add(encoded_varint_field_len(
                DRAWABLE_LOCKED_FIELD,
                u64::from(value),
            ))
            .ok_or_else(DecodeError::projection)?;
    }
    Ok(output_bytes)
}

fn emit_lock_rewrite(
    source: &[u8],
    write: TableInfoLockWrite,
    options: DecodeOptions,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let nested_options = options.descend()?;
    let mut saw_super = false;
    let mut saw_model = false;
    let mut remaining = source;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let field = next_strict_field_without_budget(&mut remaining, options)?;
        let end = source.len() - remaining.len();
        match field.number {
            TABLE_SUPER_FIELD => {
                if saw_super {
                    return Err(DecodeError::duplicate_singular(
                        "TST.TableInfoArchive.super",
                    ));
                }
                saw_super = true;
                let payload = field.length_delimited()?;
                let nested_len = measure_drawable_unbudgeted(payload, write, nested_options)?;
                append_length_delimited_field(output, TABLE_SUPER_FIELD, nested_len)?;
                emit_drawable_rewrite(payload, write, nested_options, output)?;
            },
            TABLE_MODEL_FIELD => {
                if saw_model {
                    return Err(DecodeError::duplicate_singular(
                        "TST.TableInfoArchive.table_model",
                    ));
                }
                saw_model = true;
                append_rewrite_bytes(output, &source[start..end])?;
            },
            _ => append_rewrite_bytes(output, &source[start..end])?,
        }
    }
    if !saw_super || !saw_model {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn measure_drawable_unbudgeted(
    source: &[u8],
    write: TableInfoLockWrite,
    options: DecodeOptions,
) -> Result<usize, DecodeError> {
    let mut output_bytes = 0usize;
    let mut current = None;
    let mut remaining = source;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let field = next_strict_field_without_budget(&mut remaining, options)?;
        let end = source.len() - remaining.len();
        if field.number == DRAWABLE_LOCKED_FIELD {
            if current.is_some() {
                return Err(DecodeError::duplicate_singular(
                    "TSD.DrawableArchive.locked",
                ));
            }
            current = Some(require_canonical_bool(field.varint()?)?);
        } else {
            output_bytes = output_bytes
                .checked_add(end - start)
                .ok_or_else(DecodeError::projection)?;
        }
    }
    if let Some(value) = effective_lock_value(current, write) {
        output_bytes = output_bytes
            .checked_add(encoded_varint_field_len(
                DRAWABLE_LOCKED_FIELD,
                u64::from(value),
            ))
            .ok_or_else(DecodeError::projection)?;
    }
    Ok(output_bytes)
}

fn emit_drawable_rewrite(
    source: &[u8],
    write: TableInfoLockWrite,
    options: DecodeOptions,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut current = None;
    let mut remaining = source;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let field = next_strict_field_without_budget(&mut remaining, options)?;
        let end = source.len() - remaining.len();
        if field.number == DRAWABLE_LOCKED_FIELD {
            if current.is_some() {
                return Err(DecodeError::duplicate_singular(
                    "TSD.DrawableArchive.locked",
                ));
            }
            current = Some(require_canonical_bool(field.varint()?)?);
        } else {
            append_rewrite_bytes(output, &source[start..end])?;
        }
    }
    if let Some(value) = effective_lock_value(current, write) {
        append_varint_field(output, DRAWABLE_LOCKED_FIELD, u64::from(value))?;
    }
    Ok(())
}

fn next_strict_field_without_budget<'source>(
    source: &mut &'source [u8],
    options: DecodeOptions,
) -> Result<StrictField<'source>, DecodeError> {
    let mut budget = Budget::new(DecodeOptions::new(
        source.len().max(1),
        usize::MAX,
        usize::MAX,
        options.recursion_limit,
    ));
    next_strict_field(source, options, &mut budget)?
        .ok_or_else(|| buffa::DecodeError::UnexpectedEof.into())
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn encoded_varint_field_len(number: u32, value: u64) -> usize {
    varint_len(u64::from(number) << 3) + varint_len(value)
}

fn encoded_length_delimited_field_len(
    number: u32,
    payload_len: usize,
) -> Result<usize, DecodeError> {
    let payload_len = u64::try_from(payload_len).map_err(|_error| DecodeError::projection())?;
    varint_len(u64::from(number) << 3)
        .checked_add(varint_len(payload_len))
        .and_then(|length| length.checked_add(payload_len as usize))
        .ok_or_else(DecodeError::projection)
}

fn reserve_rewrite_output(amount: usize) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(amount)
        .map_err(|_error| DecodeError::allocation(amount))?;
    Ok(output)
}

fn append_rewrite_bytes(output: &mut Vec<u8>, bytes: &[u8]) -> Result<(), DecodeError> {
    let remaining = output
        .capacity()
        .checked_sub(output.len())
        .ok_or_else(DecodeError::projection)?;
    if bytes.len() > remaining {
        return Err(DecodeError::projection());
    }
    output.extend_from_slice(bytes);
    Ok(())
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) -> Result<(), DecodeError> {
    let required = encoded_varint_field_len(number, value);
    let remaining = output
        .capacity()
        .checked_sub(output.len())
        .ok_or_else(DecodeError::projection)?;
    if required > remaining {
        return Err(DecodeError::projection());
    }
    push_varint(u64::from(number) << 3, output);
    push_varint(value, output);
    Ok(())
}

fn append_length_delimited_field(
    output: &mut Vec<u8>,
    number: u32,
    payload_len: usize,
) -> Result<(), DecodeError> {
    let required = encoded_length_delimited_field_len(number, payload_len)?;
    let remaining = output
        .capacity()
        .checked_sub(output.len())
        .ok_or_else(DecodeError::projection)?;
    if required - payload_len > remaining {
        return Err(DecodeError::projection());
    }
    push_varint((u64::from(number) << 3) | 2, output);
    push_varint(
        u64::try_from(payload_len).map_err(|_error| DecodeError::projection())?,
        output,
    );
    Ok(())
}

fn push_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn check_prepare_rewrite_limits(
    requirements: RewriteExecutionRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    check_rewrite_execution_limits(
        requirements,
        RewriteExecutionLimits {
            input_bytes: options.max_message_bytes,
            output_bytes: options.max_output_bytes,
            fields: options.max_fields,
            work_bytes: options.max_work_bytes,
            max_depth: options.recursion_limit,
            allocations: options.max_allocations,
            retained_bytes: options.max_retained_bytes,
            scratch_bytes: options.max_scratch_bytes,
        },
    )
}

fn check_rewrite_execution_limits(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    if requirements.input_bytes > limits.input_bytes {
        return Err(DecodeError::wire_resource_limit_error(
            WireResourceLimit::Bytes {
                observed: Some(requirements.input_bytes),
                maximum: Some(limits.input_bytes),
            },
        ));
    }
    if requirements.output_bytes > limits.output_bytes {
        return Err(DecodeError::output_limit(
            requirements.output_bytes,
            limits.output_bytes,
        ));
    }
    if requirements.fields > limits.fields {
        return Err(DecodeError::field_limit(requirements.fields, limits.fields));
    }
    if requirements.work_bytes > limits.work_bytes {
        return Err(DecodeError::work_limit(
            requirements.work_bytes,
            limits.work_bytes,
        ));
    }
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::wire_resource_limit_error(
            WireResourceLimit::Nesting {
                observed: Some(requirements.max_depth),
                maximum: Some(limits.max_depth),
            },
        ));
    }
    if requirements.allocations > limits.allocations {
        return Err(DecodeError::allocation_limit(
            requirements.allocations,
            limits.allocations,
        ));
    }
    if requirements.retained_bytes > limits.retained_bytes {
        return Err(DecodeError::retained_limit(
            requirements.retained_bytes,
            limits.retained_bytes,
        ));
    }
    if requirements.scratch_bytes > limits.scratch_bytes {
        return Err(DecodeError::scratch_limit(
            requirements.scratch_bytes,
            limits.scratch_bytes,
        ));
    }
    Ok(())
}

fn require_canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}

#[derive(Clone, Copy, Debug)]
enum StrictValue<'source> {
    Varint(u64),
    Fixed64,
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32,
}

#[derive(Clone, Copy, Debug)]
struct StrictField<'source> {
    number: u32,
    wire_type: buffa::encoding::WireType,
    value: StrictValue<'source>,
}

impl<'source> StrictField<'source> {
    fn require_wire_type(self, expected: buffa::encoding::WireType) -> Result<(), DecodeError> {
        if self.wire_type != expected {
            return Err(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected: expected as u8,
                actual: self.wire_type as u8,
            }
            .into());
        }
        Ok(())
    }

    fn varint(self) -> Result<u64, DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::Varint)?;
        match self.value {
            StrictValue::Varint(value) => Ok(value),
            StrictValue::Fixed64
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group
            | StrictValue::Fixed32 => Err(DecodeError::projection()),
        }
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::LengthDelimited)?;
        match self.value {
            StrictValue::LengthDelimited(value) => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64
            | StrictValue::Group
            | StrictValue::Fixed32 => Err(DecodeError::projection()),
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum ParseItem<'source> {
    Field(StrictField<'source>),
    EndGroup(u32),
}

fn next_strict_field<'source>(
    source: &mut &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_strict_field(source, options.recursion_limit, budget)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(number)) => {
            Err(buffa::DecodeError::InvalidEndGroup(number).into())
        },
        None => Ok(None),
    }
}

fn parse_strict_field<'source>(
    source: &mut &'source [u8],
    recursion_limit: u32,
    budget: &mut Budget,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    let (encoded_tag, canonical_key) = take_varint(source)?;
    if !canonical_key {
        return Err(DecodeError::noncanonical("protobuf field key"));
    }
    budget.charge_field()?;
    let raw_tag =
        u32::try_from(encoded_tag).map_err(|_conversion| buffa::DecodeError::InvalidFieldNumber)?;
    let field_number = raw_tag >> 3;
    if field_number == 0 || field_number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let raw_wire_type = raw_tag & 7;
    let wire_type = buffa::encoding::WireType::from_u32(raw_wire_type)?;
    let value = match wire_type {
        buffa::encoding::WireType::Varint => {
            let (value, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            }
            StrictValue::Varint(value)
        },
        buffa::encoding::WireType::Fixed64 => {
            let _bytes = take_exact(source, 8)?;
            StrictValue::Fixed64
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (encoded_length, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("length-delimited size"));
            }
            let length = usize::try_from(encoded_length)
                .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
            StrictValue::LengthDelimited(take_exact(source, length)?)
        },
        buffa::encoding::WireType::StartGroup => {
            let child_limit = recursion_limit
                .checked_sub(1)
                .ok_or_else(DecodeError::recursion_limit)?;
            skip_strict_group(source, field_number, child_limit, budget)?;
            StrictValue::Group
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            let _bytes = take_exact(source, 4)?;
            StrictValue::Fixed32
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw_wire_type).into()),
    };
    Ok(Some(ParseItem::Field(StrictField {
        number: field_number,
        wire_type,
        value,
    })))
}

fn skip_strict_group(
    source: &mut &[u8],
    expected_field_number: u32,
    recursion_limit: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    loop {
        match parse_strict_field(source, recursion_limit, budget)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected_field_number => return Ok(()),
            Some(ParseItem::EndGroup(number)) => {
                return Err(buffa::DecodeError::InvalidEndGroup(number).into());
            },
            None => return Err(buffa::DecodeError::UnexpectedEof.into()),
        }
    }
}

fn take_varint(source: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original
            .get(index)
            .ok_or(buffa::DecodeError::UnexpectedEof)?;
        if index == 9 && byte > 1 {
            return Err(buffa::DecodeError::VarintTooLong.into());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            *source = &original[consumed..];
            return Ok((value, canonical_varint_len(value) == consumed));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}

fn canonical_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn take_exact<'source>(
    source: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], DecodeError> {
    if source.len() < length {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    let (selected, remaining) = source.split_at(length);
    *source = remaining;
    Ok(selected)
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::shadow_unrelated,
    reason = "Focused negative tests use explicit panic messages and reuse local error roles."
)]
mod tests {
    use std::num::NonZeroU64;

    use prost::Message as _;

    use super::{
        Budget, DecodeOptions, RewriteExecutionLimits, TableInfoLockWrite, TableInfoSnapshot,
        TableModelReference, WireResourceLimit, decode_table_info, decode_table_model_reference,
        prepare_table_info_lock_rewrite, prepare_table_info_lock_rewrite_with_fingerprint,
        rewrite_table_info_lock, table_info_source_fingerprint,
    };
    use crate::{tsd, tsp, tst};

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            source.len().max(1),
            source.len().saturating_mul(4).max(1),
            2,
        )
    }

    fn decode(source: &[u8]) -> Result<TableModelReference, super::DecodeError> {
        decode_table_model_reference(source, options(source))
    }

    fn decode_snapshot(source: &[u8]) -> Result<TableInfoSnapshot, super::DecodeError> {
        decode_table_info(source, options(source))
    }

    fn table_model_field(model: &[u8]) -> Vec<u8> {
        let mut output = vec![0x12, u8::try_from(model.len()).expect("small test payload")];
        output.extend_from_slice(model);
        output
    }

    fn table_info(model: &[u8]) -> Vec<u8> {
        [vec![0x0a, 0x00], table_model_field(model)].concat()
    }

    fn table_info_with_super(super_payload: &[u8]) -> Vec<u8> {
        let mut source = vec![
            0x0a,
            u8::try_from(super_payload.len()).expect("small test payload"),
        ];
        source.extend_from_slice(super_payload);
        source.extend(table_model_field(&[0x08, 0x2a]));
        source
    }

    fn rewrite_options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
            .with_max_message_bytes(source.len())
            .with_max_fields(source.len().saturating_mul(32).max(32))
            .with_max_work_bytes(source.len().saturating_mul(128).max(128))
            .with_max_output_bytes(source.len().saturating_add(64))
            .with_max_scratch_bytes(source.len().saturating_add(128))
    }

    #[test]
    fn canonical_prost_table_info_matches_the_strict_projection()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = tst::TableInfoArchive {
            super_: tsd::DrawableArchive::default(),
            table_model: tsp::Reference {
                identifier: 42,
                deprecated_type: Some(-7),
                deprecated_is_external: Some(false),
            },
            ..tst::TableInfoArchive::default()
        }
        .encode_to_vec();

        assert_eq!(
            decode(&source)?.identifier(),
            NonZeroU64::new(42).expect("non-zero test identifier")
        );
        Ok(())
    }

    #[test]
    fn prost_snapshot_matches_buffa_for_absent_false_and_true_drawable_locks()
    -> Result<(), Box<dyn std::error::Error>> {
        for locked in [None, Some(false), Some(true)] {
            let source = tst::TableInfoArchive {
                super_: tsd::DrawableArchive {
                    locked,
                    ..Default::default()
                },
                table_model: tsp::Reference {
                    identifier: 42,
                    ..Default::default()
                },
                ..Default::default()
            }
            .encode_to_vec();
            let snapshot = decode_snapshot(&source)?;
            assert_eq!(
                snapshot.table_model().identifier(),
                NonZeroU64::new(42).expect("non-zero test identifier")
            );
            assert_eq!(snapshot.locked(), locked);
        }
        Ok(())
    }

    #[test]
    fn unselected_drawable_and_table_info_metadata_remain_opaque()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = [
            0x0a, 0x02, 0x18, 0x7f, // unselected DrawableArchive metadata
            0x12, 0x02, 0x08, 0x2a, // selected table-model reference
            0x1a, 0x01, 0xff, // opaque editing_state metadata
        ];
        assert_eq!(
            decode(&source)?.identifier(),
            NonZeroU64::new(42).expect("non-zero test identifier")
        );
        Ok(())
    }

    #[test]
    fn required_and_unique_table_model_identifier_are_enforced() {
        let error = decode(&[]).expect_err("missing drawable super envelope");
        assert_eq!(
            error.missing_required_field(),
            Some("TST.TableInfoArchive.super")
        );

        let error = decode(&[0x0a, 0x00]).expect_err("missing model reference");
        assert_eq!(
            error.missing_required_field(),
            Some("TST.TableInfoArchive.table_model")
        );

        let error = decode(&table_info(&[])).expect_err("missing nested identifier");
        assert_eq!(
            error.missing_required_field(),
            Some("TSP.Reference.identifier")
        );

        let duplicate_model = [
            vec![0x0a, 0x00],
            table_model_field(&[0x08, 0x01]),
            table_model_field(&[0x08, 0x02]),
        ]
        .concat();
        let error = decode(&duplicate_model).expect_err("duplicate model reference");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TST.TableInfoArchive.table_model")
        );

        let error = decode(&table_info(&[0x08, 0x01, 0x08, 0x02]))
            .expect_err("duplicate nested identifier");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSP.Reference.identifier")
        );

        let duplicate_super = [
            vec![0x0a, 0x00],
            vec![0x0a, 0x00],
            table_model_field(&[0x08, 0x01]),
        ]
        .concat();
        let error = decode(&duplicate_super).expect_err("duplicate drawable super envelope");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TST.TableInfoArchive.super")
        );
    }

    #[test]
    fn wrong_wire_zero_and_malformed_selected_fields_are_rejected() {
        let root_wrong_wire = [0x0a, 0x00, 0x10, 0x2a];
        assert!(decode(&root_wrong_wire).is_err());

        let super_wrong_wire = [0x08, 0x00, 0x12, 0x02, 0x08, 0x2a];
        assert!(decode(&super_wrong_wire).is_err());

        let nested_wrong_wire = table_info(&[0x0a, 0x00]);
        assert!(decode(&nested_wrong_wire).is_err());

        let error = decode(&table_info(&[0x08, 0x00])).expect_err("zero identifier");
        assert_eq!(
            error.zero_identifier_field(),
            Some("TSP.Reference.identifier")
        );

        let malformed = [0x0a, 0x00, 0x12, 0x01, 0x08];
        assert!(decode(&malformed).is_err());

        let duplicate_lock = [
            0x0a, 0x04, 0x28, 0x00, 0x28, 0x01, // duplicate DrawableArchive.locked
            0x12, 0x02, 0x08, 0x2a,
        ];
        assert_eq!(
            decode(&duplicate_lock)
                .expect_err("duplicate drawable lock")
                .duplicate_singular_field(),
            Some("TSD.DrawableArchive.locked")
        );

        let wrong_lock_wire = [0x0a, 0x02, 0x2a, 0x00, 0x12, 0x02, 0x08, 0x2a];
        assert!(decode(&wrong_lock_wire).is_err());

        let nonboolean_lock = [0x0a, 0x02, 0x28, 0x02, 0x12, 0x02, 0x08, 0x2a];
        assert_eq!(
            decode(&nonboolean_lock)
                .expect_err("non-boolean drawable lock")
                .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
    }

    #[test]
    fn noncanonical_selected_and_opaque_framing_is_rejected() {
        let overlong_super_length = [0x0a, 0x80, 0x00, 0x12, 0x02, 0x08, 0x2a];
        assert_eq!(
            decode(&overlong_super_length)
                .expect_err("overlong drawable super length")
                .noncanonical_reason(),
            Some("length-delimited size")
        );

        let overlong_key = [0x0a, 0x00, 0x92, 0x00, 0x02, 0x08, 0x2a];
        assert_eq!(
            decode(&overlong_key)
                .expect_err("overlong root key")
                .noncanonical_reason(),
            Some("protobuf field key")
        );

        let overlong_length = [0x0a, 0x00, 0x12, 0x82, 0x00, 0x08, 0x2a];
        assert_eq!(
            decode(&overlong_length)
                .expect_err("overlong nested length")
                .noncanonical_reason(),
            Some("length-delimited size")
        );

        let overlong_identifier = table_info(&[0x08, 0xaa, 0x00]);
        assert_eq!(
            decode(&overlong_identifier)
                .expect_err("overlong identifier")
                .noncanonical_reason(),
            Some("protobuf varint value")
        );

        let noncanonical_opaque = [0x0a, 0x00, 0x18, 0x81, 0x00, 0x12, 0x02, 0x08, 0x2a];
        assert_eq!(
            decode(&noncanonical_opaque)
                .expect_err("noncanonical opaque metadata")
                .noncanonical_reason(),
            Some("protobuf varint value")
        );

        let noncanonical_lock = [0x0a, 0x03, 0x28, 0x81, 0x00, 0x12, 0x02, 0x08, 0x2a];
        assert_eq!(
            decode(&noncanonical_lock)
                .expect_err("noncanonical drawable lock")
                .noncanonical_reason(),
            Some("protobuf varint value")
        );
    }

    #[test]
    fn canonical_unknown_wire_forms_remain_opaque_and_source_owned() {
        let source = [
            0x0a, 0x16, // DrawableArchive.super
            0x08, 0x07, // unknown varint
            0x11, 0, 1, 2, 3, 4, 5, 6, 7, // unknown fixed64
            0x1a, 0x02, 0xaa, 0xbb, // unknown length-delimited bytes
            0x25, 9, 8, 7, 6, // unknown fixed32
            0x28, 0x01, // selected locked=true
            0x12, 0x08, // TableModelReference
            0x08, 0x2a, // selected identifier=42
            0x10, 0x09, // unknown varint in the nested reference
            0x1b, 0x08, 0x0b, 0x1c, // unknown group in the nested reference
            0x98, 0x06, 0x01, // root unknown varint
            0xa1, 0x06, 0, 1, 2, 3, 4, 5, 6, 7, // root unknown fixed64
            0xaa, 0x06, 0x01, 0xff, // root unknown bytes
            0xb5, 0x06, 0, 1, 2, 3, // root unknown fixed32
            0xbb, 0x06, 0x08, 0x01, 0xbc, 0x06, // root unknown group
        ];
        let before = source;

        let snapshot = decode_table_info(
            &source,
            DecodeOptions::new(source.len(), 32, source.len() * 4, 3),
        )
        .expect("canonical unknown fields remain opaque");
        assert_eq!(
            snapshot.table_model().identifier(),
            NonZeroU64::new(42).expect("non-zero test identifier")
        );
        assert_eq!(snapshot.locked(), Some(true));
        assert_eq!(
            source, before,
            "projection must not rewrite caller-owned bytes"
        );
    }

    #[test]
    fn malformed_unknown_groups_and_nested_noncanonical_values_fail_closed() {
        let malformed_groups: [&[u8]; 3] = [
            &[0x0a, 0x01, 0x0b, 0x12, 0x02, 0x08, 0x01],
            &[0x0a, 0x03, 0x0b, 0x08, 0x01, 0x12, 0x02, 0x08, 0x01],
            &[0x0a, 0x03, 0x0b, 0x0c, 0x0c, 0x12, 0x02, 0x08, 0x01],
        ];
        for source in malformed_groups {
            assert!(
                decode(&source).is_err(),
                "malformed unknown group: {source:?}"
            );
        }

        let nested_noncanonical = [
            0x0a, 0x03, 0x08, 0x81, 0x00, // noncanonical unknown nested varint
            0x12, 0x02, 0x08, 0x01,
        ];
        assert_eq!(
            decode(&nested_noncanonical)
                .expect_err("noncanonical nested unknown value")
                .noncanonical_reason(),
            Some("protobuf varint value")
        );
    }

    #[test]
    fn nested_unknown_fields_consume_field_and_work_budgets()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = [
            0x0a, 0x02, 0x08, 0x01, // one nested unknown field in super
            0x12, 0x04, 0x08, 0x2a, 0x10, 0x09, // identifier + nested unknown
        ];
        // Root has two fields and the selected nested messages have three
        // fields total, for five strict visits. Message work is charged for
        // the root and both selected nested payloads: 2 * (10 + 2 + 4).
        let exact = DecodeOptions::new(source.len(), 5, 32, 2);
        assert_eq!(
            decode_table_model_reference(&source, exact)?
                .identifier()
                .get(),
            42
        );

        assert_eq!(
            decode_table_model_reference(&source, DecodeOptions::new(source.len(), 4, 32, 2))
                .expect_err("nested field cap")
                .field_limit_values(),
            Some((5, 4))
        );
        assert_eq!(
            decode_table_model_reference(&source, DecodeOptions::new(source.len(), 5, 31, 2))
                .expect_err("nested work cap")
                .work_limit_values(),
            Some((32, 31))
        );
        Ok(())
    }

    #[test]
    fn exact_boundary_limits_are_accepted_and_one_less_is_rejected()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = [0x0a, 0x00, 0x12, 0x02, 0x08, 0x01];
        let exact = DecodeOptions::new(source.len(), 3, 16, 2);
        assert_eq!(
            decode_table_model_reference(&source, exact)?.identifier(),
            NonZeroU64::new(1).expect("non-zero test identifier")
        );
        assert_eq!(
            decode_table_model_reference(&source, DecodeOptions::new(source.len() - 1, 3, 16, 2))
                .expect_err("byte cap")
                .wire_resource_limit(),
            Some(WireResourceLimit::Bytes {
                observed: Some(source.len()),
                maximum: Some(source.len() - 1),
            })
        );
        assert_eq!(
            decode_table_model_reference(&source, DecodeOptions::new(source.len(), 2, 16, 2))
                .expect_err("field cap")
                .field_limit_values(),
            Some((3, 2))
        );
        assert_eq!(
            decode_table_model_reference(&source, DecodeOptions::new(source.len(), 3, 15, 2))
                .expect_err("work cap")
                .work_limit_values(),
            Some((16, 15))
        );
        assert_eq!(
            decode_table_model_reference(&source, DecodeOptions::new(source.len(), 3, 16, 0))
                .expect_err("zero nesting cap")
                .wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: Some(0),
                maximum: Some(64),
            })
        );
        Ok(())
    }

    #[test]
    fn budget_charge_overflow_is_a_typed_limit() {
        let mut budget = Budget::new(DecodeOptions::new(1, usize::MAX, 1, 2));
        budget.fields = usize::MAX;
        let error = budget
            .charge_field()
            .expect_err("field addition must not saturate");
        assert_eq!(error.field_limit_values(), Some((usize::MAX, usize::MAX)));

        let mut budget = Budget::new(DecodeOptions::new(1, 1, usize::MAX, 2));
        budget.work_bytes = usize::MAX - 1;
        let error = budget
            .charge_message(1)
            .expect_err("work addition must not saturate");
        assert_eq!(error.work_limit_values(), Some((usize::MAX, usize::MAX)));

        let mut budget = Budget::new(DecodeOptions::new(1, 1, usize::MAX, 2));
        let error = budget
            .charge_message(usize::MAX)
            .expect_err("work multiplication must not saturate");
        assert_eq!(error.work_limit_values(), Some((usize::MAX, usize::MAX)));
    }

    #[test]
    fn exhausted_nested_wire_depth_has_a_typed_error() {
        let source = [
            0x0a, 0x02, 0x0b, 0x0c, // one unknown group in DrawableArchive
            0x12, 0x02, 0x08, 0x01,
        ];
        assert_eq!(
            decode_table_info(
                &source,
                DecodeOptions::new(source.len(), source.len(), source.len() * 4, 1),
            )
            .expect_err("nested wire depth")
            .wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: None,
                maximum: Some(1),
            })
        );
    }

    #[test]
    fn structural_wire_failures_never_panic() {
        let malformed: [&[u8]; 6] = [
            &[0x80],
            &[0x00],
            &[0x0f],
            &[0x12, 0x02, 0x08],
            &[0x0b],
            &[
                0x08, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x02,
            ],
        ];
        for source in malformed {
            assert!(decode(source).is_err());
        }
    }

    #[test]
    fn prepared_lock_rewrite_preserves_absent_false_and_true_states()
    -> Result<(), Box<dyn std::error::Error>> {
        let absent = table_info(&[0x08, 0x2a]);
        let explicit_false = table_info_with_super(&[0x28, 0x00]);
        let explicit_true = table_info_with_super(&[0x28, 0x01]);

        // Semantic unlocked writes preserve a missing lock field, while an
        // explicit source false remains explicit.  The changed directions
        // use only the canonical one-byte bool representation.
        let cases = [
            (&absent, false, false),
            (&absent, true, true),
            (&explicit_false, false, false),
            (&explicit_false, true, true),
            (&explicit_true, false, true),
            (&explicit_true, true, false),
        ];
        for (source, locked, changed) in cases {
            let options = rewrite_options(source);
            let prepared =
                prepare_table_info_lock_rewrite(source, TableInfoLockWrite::new(locked), options)?;
            assert_eq!(prepared.prepare_report().changed(), changed);
            let output = prepared
                .execute(prepared.execution_requirements().exact())?
                .into_bytes();
            let snapshot = decode_snapshot(&output)?;
            let expected = if source.as_slice() == absent.as_slice() && !locked {
                None
            } else {
                Some(locked)
            };
            assert_eq!(snapshot.locked(), expected);
            if !changed {
                assert_eq!(output, *source);
            }
        }

        let removed = rewrite_table_info_lock(
            &explicit_false,
            TableInfoLockWrite::explicit(None),
            rewrite_options(&explicit_false),
        )?;
        assert_eq!(decode_snapshot(&removed)?.locked(), None);
        Ok(())
    }

    #[test]
    fn prepared_lock_rewrite_retains_unknown_framing_and_supports_fingerprint()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = [
            0x0a, 0x08, // DrawableArchive.super
            0x08, 0x07, // unknown varint
            0x1b, 0x08, 0x0b, 0x1c, // balanced unknown group
            0x28, 0x01, // selected locked=true
            0x12, 0x02, 0x08, 0x2a, // required table model reference
            0x1a, 0x01, 0xff, // opaque root bytes
            0x2d, 0x01, 0x02, 0x03, 0x04, // opaque root fixed32
        ];
        let fingerprint = table_info_source_fingerprint(&source);
        let prepared = prepare_table_info_lock_rewrite_with_fingerprint(
            &source,
            fingerprint,
            TableInfoLockWrite::new(false),
            rewrite_options(&source),
        )?;
        assert_eq!(prepared.source_fingerprint(), fingerprint);
        let output = prepared
            .execute(prepared.execution_requirements().exact())?
            .into_bytes();
        assert!(
            output
                .windows(6)
                .any(|window| { window == [0x08, 0x07, 0x1b, 0x08, 0x0b, 0x1c] })
        );
        assert!(output.windows(3).any(|window| window == [0x1a, 0x01, 0xff]));
        assert_eq!(decode_snapshot(&output)?.locked(), Some(false));

        let stale = prepare_table_info_lock_rewrite(
            &source,
            TableInfoLockWrite::new(true).expecting_fingerprint(fingerprint ^ 1),
            rewrite_options(&source),
        )
        .expect_err("stale source fingerprint");
        assert!(stale.is_fingerprint_mismatch());
        Ok(())
    }

    #[test]
    fn prepared_lock_rewrite_rejects_malformed_sources_before_output()
    -> Result<(), Box<dyn std::error::Error>> {
        let malformed: [&[u8]; 5] = [
            // duplicate selected lock field
            &[0x0a, 0x04, 0x28, 0x00, 0x28, 0x01, 0x12, 0x02, 0x08, 0x2a],
            // selected lock has the wrong wire type
            &[0x0a, 0x02, 0x2a, 0x00, 0x12, 0x02, 0x08, 0x2a],
            // selected lock has a non-canonical varint value
            &[0x0a, 0x03, 0x28, 0x81, 0x00, 0x12, 0x02, 0x08, 0x2a],
            // unbalanced unknown group inside the selected envelope
            &[0x0a, 0x01, 0x0b, 0x12, 0x02, 0x08, 0x2a],
            // missing required model reference
            &[0x0a, 0x00],
        ];
        for source in malformed {
            assert!(
                prepare_table_info_lock_rewrite(
                    source,
                    TableInfoLockWrite::new(true),
                    rewrite_options(source),
                )
                .is_err(),
                "malformed source was admitted: {source:?}"
            );
        }
        Ok(())
    }

    #[test]
    fn prepared_lock_rewrite_enforces_each_execution_limit_at_max_minus_one()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = table_info(&[0x08, 0x2a]);
        let options = rewrite_options(&source);
        let baseline =
            prepare_table_info_lock_rewrite(&source, TableInfoLockWrite::new(true), options)?;
        let requirements = baseline.execution_requirements();
        assert!(requirements.input_bytes() > 0);
        assert!(requirements.output_bytes() > 0);
        assert!(requirements.fields() > 0);
        assert!(requirements.work_bytes() > 0);
        assert!(requirements.max_depth() > 0);
        assert!(requirements.allocations() > 0);
        assert!(requirements.retained_bytes() > 0);
        assert!(requirements.scratch_bytes() > 0);

        let input_error =
            prepare_table_info_lock_rewrite(&source, TableInfoLockWrite::new(true), options)?
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_input_bytes(requirements.input_bytes() - 1),
                )
                .expect_err("input ceiling");
        assert!(input_error.wire_resource_limit().is_some());

        let output_error =
            prepare_table_info_lock_rewrite(&source, TableInfoLockWrite::new(true), options)?
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_output_bytes(requirements.output_bytes() - 1),
                )
                .expect_err("output ceiling");
        assert_eq!(
            output_error.output_limit_values(),
            Some((requirements.output_bytes(), requirements.output_bytes() - 1))
        );

        let field_error =
            prepare_table_info_lock_rewrite(&source, TableInfoLockWrite::new(true), options)?
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_fields(requirements.fields() - 1),
                )
                .expect_err("field ceiling");
        assert!(field_error.field_limit_values().is_some());

        let work_error =
            prepare_table_info_lock_rewrite(&source, TableInfoLockWrite::new(true), options)?
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_work_bytes(requirements.work_bytes() - 1),
                )
                .expect_err("work ceiling");
        assert!(work_error.work_limit_values().is_some());

        let nesting_error =
            prepare_table_info_lock_rewrite(&source, TableInfoLockWrite::new(true), options)?
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_max_depth(requirements.max_depth() - 1),
                )
                .expect_err("nesting ceiling");
        assert!(nesting_error.wire_resource_limit().is_some());

        let allocation_error =
            prepare_table_info_lock_rewrite(&source, TableInfoLockWrite::new(true), options)?
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_allocations(requirements.allocations() - 1),
                )
                .expect_err("allocation ceiling");
        assert!(allocation_error.allocation_limit_values().is_some());

        let retained_error =
            prepare_table_info_lock_rewrite(&source, TableInfoLockWrite::new(true), options)?
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_retained_bytes(requirements.retained_bytes() - 1),
                )
                .expect_err("retained ceiling");
        assert!(retained_error.retained_limit_values().is_some());

        let scratch_error = baseline
            .execute(
                RewriteExecutionLimits::exact(requirements)
                    .with_scratch_bytes(requirements.scratch_bytes() - 1),
            )
            .expect_err("scratch ceiling");
        assert!(scratch_error.scratch_limit_values().is_some());
        Ok(())
    }
}
