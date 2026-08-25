//! Strict, borrowed routing for `TP.SectionTemplateArchive` header/footer
//! references.
//!
//! The native envelope has two repeated fields of complete `TSP.Reference`
//! messages (`headers = 1` and `footers = 2`).  This codec deliberately keeps
//! those repeated records on the caller-owned wire.  It validates the selected
//! known fields, cross-checks each reference with the private Buffa lazy view,
//! and never asks a generated type to own or encode the repeated records.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the strict wire reader is intentionally arranged before its low-level helpers"
)]

use core::fmt;

#[cfg(test)]
use std::cell::Cell;

use buffa::DecodeOptions as BuffaDecodeOptions;

// The reference projection is schema-identical to TSP.Reference and is
// already generated in isolation for the Pages drawable-order codec.  Reusing
// it avoids a second generated closure and keeps this envelope handwritten.
use crate::buffa_pages_drawable_order_generated::LitchiIwaProjection as projection;

const HEADERS_FIELD: u32 = 1;
const FOOTERS_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const ROOT_DEPTH: u32 = 1;
const REFERENCE_DEPTH: u32 = 2;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Finite limits for one complete section-template payload or rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_references: usize,
}

impl DecodeOptions {
    /// Construct an explicit finite input/output, field, work, depth, and
    /// repeated-reference policy.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_references: usize,
    ) -> Self {
        Self {
            max_input_bytes,
            max_output_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_references,
        }
    }

    /// Build a conservative finite profile from one source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(2).max(1),
            bytes.saturating_mul(8).max(1),
            bytes.saturating_mul(32).max(1),
            8,
            bytes.max(1),
        )
    }

    /// Replace the candidate output ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the repeated-reference ceiling.
    #[must_use]
    pub const fn with_max_references(mut self, maximum: usize) -> Self {
        self.max_references = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_input_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Borrowed scalar facts from one complete `TSP.Reference`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReferenceSnapshot<'source> {
    raw: &'source [u8],
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

impl<'source> ReferenceSnapshot<'source> {
    /// Required nonzero native identifier.
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }

    /// Optional deprecated reference type, preserving presence.
    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }

    /// Optional deprecated external marker, preserving presence.
    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }

    /// Exact nested reference bytes borrowed from the source payload.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

/// Borrowed complete section-template projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionTemplateSnapshot<'source> {
    source: &'source [u8],
    headers: usize,
    footers: usize,
}

impl<'source> SectionTemplateSnapshot<'source> {
    /// Exact source payload bytes retained by the snapshot contract.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }

    /// Number of repeated header references.
    #[must_use]
    pub const fn header_count(self) -> usize {
        self.headers
    }

    /// Number of repeated footer references.
    #[must_use]
    pub const fn footer_count(self) -> usize {
        self.footers
    }

    /// Alias for [`Self::header_count`].
    #[must_use]
    pub const fn headers_len(self) -> usize {
        self.headers
    }

    /// Alias for [`Self::footer_count`].
    #[must_use]
    pub const fn footers_len(self) -> usize {
        self.footers
    }

    /// Whether the envelope has neither header nor footer reference.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.headers == 0 && self.footers == 0
    }

    /// Borrow complete header references in source order.
    #[must_use]
    pub fn headers(self) -> ReferenceIter<'source> {
        ReferenceIter::new(self.source, HEADERS_FIELD, self.headers)
    }

    /// Borrow complete footer references in source order.
    #[must_use]
    pub fn footers(self) -> ReferenceIter<'source> {
        ReferenceIter::new(self.source, FOOTERS_FIELD, self.footers)
    }

    /// Alias for [`Self::headers`].
    #[must_use]
    pub fn header_references(self) -> ReferenceIter<'source> {
        self.headers()
    }

    /// Alias for [`Self::footers`].
    #[must_use]
    pub fn footer_references(self) -> ReferenceIter<'source> {
        self.footers()
    }

    /// Borrow only header identifiers without allocating.
    #[must_use]
    pub fn header_identifiers(self) -> IdentifierIter<'source> {
        IdentifierIter::new(self.headers())
    }

    /// Borrow only footer identifiers without allocating.
    #[must_use]
    pub fn footer_identifiers(self) -> IdentifierIter<'source> {
        IdentifierIter::new(self.footers())
    }
}

/// Allocation-free iterator over one repeated role's complete references.
#[derive(Debug, Clone, Copy)]
pub struct ReferenceIter<'source> {
    remaining: &'source [u8],
    field: u32,
    yielded: usize,
    maximum: usize,
}

impl<'source> ReferenceIter<'source> {
    const fn new(source: &'source [u8], field: u32, maximum: usize) -> Self {
        Self {
            remaining: source,
            field,
            yielded: 0,
            maximum,
        }
    }
}

impl<'source> Iterator for ReferenceIter<'source> {
    type Item = ReferenceSnapshot<'source>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.yielded >= self.maximum {
            return None;
        }
        while !self.remaining.is_empty() {
            let field = parse_unmetered_field(&mut self.remaining).ok()??;
            if field.number != self.field {
                continue;
            }
            let UnmeteredValue::LengthDelimited(payload) = field.value else {
                return None;
            };
            let reference = parse_unmetered_reference(payload).ok()?;
            self.yielded = self.yielded.saturating_add(1);
            return Some(reference);
        }
        None
    }
}

/// Allocation-free iterator over role identifiers.
#[derive(Debug, Clone, Copy)]
pub struct IdentifierIter<'source> {
    references: ReferenceIter<'source>,
}

impl<'source> IdentifierIter<'source> {
    const fn new(references: ReferenceIter<'source>) -> Self {
        Self { references }
    }
}

impl<'source> Iterator for IdentifierIter<'source> {
    type Item = u64;

    fn next(&mut self) -> Option<Self::Item> {
        self.references.next().map(ReferenceSnapshot::identifier)
    }
}

/// A raw, already validated replacement reference.
///
/// The bytes are borrowed so unknown nested fields, source ordering, groups,
/// and overlong unknown scalar values remain the caller's preservation
/// authority.  A replacement does not canonicalize or encode a generated
/// reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReferenceWrite<'source> {
    raw: &'source [u8],
}

impl<'source> ReferenceWrite<'source> {
    /// Borrow a raw reference payload. It is validated during rewrite.
    #[must_use]
    pub const fn new(raw: &'source [u8]) -> Self {
        Self { raw }
    }

    /// Return the exact replacement bytes.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

/// Exact repeated-list replacement request for a section template.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionTemplateWrite<'source> {
    headers: &'source [ReferenceWrite<'source>],
    footers: &'source [ReferenceWrite<'source>],
}

impl<'source> SectionTemplateWrite<'source> {
    /// Construct an exact header/footer replacement request.
    #[must_use]
    pub const fn new(
        headers: &'source [ReferenceWrite<'source>],
        footers: &'source [ReferenceWrite<'source>],
    ) -> Self {
        Self { headers, footers }
    }

    /// Requested header records.
    #[must_use]
    pub const fn headers(self) -> &'source [ReferenceWrite<'source>] {
        self.headers
    }

    /// Requested footer records.
    #[must_use]
    pub const fn footers(self) -> &'source [ReferenceWrite<'source>] {
        self.footers
    }
}

/// Exact consumption report for one strict decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl DecodeReport {
    /// Source payload bytes inspected.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Aggregate strict field visits.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate bounded traversal work.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum protobuf nesting observed.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Repeated references scanned.
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    /// Handwritten allocations performed by decode.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Bytes retained by the returned snapshot contract.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Temporary scratch bytes allocated by the codec.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Exact consumption report for one source-authoritative rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    changed: bool,
}

impl RewriteReport {
    /// Source payload bytes inspected.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Exact candidate payload bytes.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Aggregate strict field visits across rewrite phases.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate bounded rewrite work.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum nesting observed across rewrite phases.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Aggregate repeated-reference visits.
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    /// Candidate output allocation events.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Bytes retained by the candidate.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Temporary scratch bytes.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Whether the candidate differs from the source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// Typed finite resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    /// Input protobuf payload exceeded its ceiling.
    InputBytes { observed: usize, maximum: usize },
    /// Candidate protobuf payload exceeded its ceiling.
    OutputBytes { observed: usize, maximum: usize },
    /// Strict field visits exceeded their ceiling.
    Fields { observed: usize, maximum: usize },
    /// Traversal work exceeded its ceiling.
    WorkBytes { observed: usize, maximum: usize },
    /// Protobuf nesting exceeded its ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// Repeated references exceeded their ceiling.
    References { observed: usize, maximum: usize },
}

/// Strict wire, semantic, or finite-resource failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Limit(WireResourceLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    Allocation { amount: usize },
    Projection,
}

impl DecodeError {
    /// Return a typed finite-limit failure, when applicable.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<WireResourceLimit> {
        match self.kind {
            DecodeErrorKind::Limit(limit) => Some(limit),
            _ => None,
        }
    }

    /// Return a failed output reservation amount, when applicable.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        match self.kind {
            DecodeErrorKind::Allocation { amount } => Some(amount),
            _ => None,
        }
    }

    /// Return the duplicate known field, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            _ => None,
        }
    }

    /// Return a stable canonicality reason, when applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            _ => None,
        }
    }

    const fn limit(limit: WireResourceLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Limit(limit),
        }
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

    const fn allocation(amount: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation { amount },
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Limit(WireResourceLimit::InputBytes { observed, maximum }) => write!(
                formatter,
                "Pages section-template input is {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(WireResourceLimit::OutputBytes { observed, maximum }) => write!(
                formatter,
                "Pages section-template output is {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(WireResourceLimit::Fields { observed, maximum }) => write!(
                formatter,
                "Pages section-template visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(WireResourceLimit::WorkBytes { observed, maximum }) => write!(
                formatter,
                "Pages section-template requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(WireResourceLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Pages section-template reached nesting {observed}; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(WireResourceLimit::References { observed, maximum }) => write!(
                formatter,
                "Pages section-template visited {observed} references; maximum is {maximum}"
            ),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => write!(formatter, "duplicate {field}"),
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::Allocation { amount } => {
                write!(
                    formatter,
                    "cannot allocate {amount} section-template output bytes"
                )
            },
            DecodeErrorKind::Projection => {
                formatter.write_str("section-template strict projection disagrees with Buffa")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self {
            kind: DecodeErrorKind::Wire(error),
        }
    }
}

/// Decode a complete `TP.SectionTemplateArchive` payload.
pub fn decode_section_template(
    source: &[u8],
    options: DecodeOptions,
) -> Result<SectionTemplateSnapshot<'_>, DecodeError> {
    Ok(decode_section_template_with_report(source, options)?.0)
}

/// Decode a complete section-template payload with exact resource accounting.
pub fn decode_section_template_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(SectionTemplateSnapshot<'_>, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let shape = scan_template(source, options, &mut budget)?;
    Ok((
        SectionTemplateSnapshot {
            source,
            headers: shape.headers,
            footers: shape.footers,
        },
        budget.decode_report(),
    ))
}

/// Compatibility alias using the native archive name.
pub fn decode_section_template_archive(
    source: &[u8],
    options: DecodeOptions,
) -> Result<SectionTemplateSnapshot<'_>, DecodeError> {
    decode_section_template(source, options)
}

/// Rewrite both repeated reference lists while preserving every unrelated raw
/// source field.  The request must contain exactly one replacement for each
/// source header/footer occurrence; it may use borrowed records from another
/// source, but each record is strictly validated before any output allocation.
pub fn rewrite_section_template(
    source: &[u8],
    write: SectionTemplateWrite<'_>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    Ok(rewrite_section_template_with_report(source, write, options)?.0)
}

/// Rewrite a section-template payload with exact output/report accounting.
pub fn rewrite_section_template_with_report(
    source: &[u8],
    write: SectionTemplateWrite<'_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let source_shape = scan_template(source, options, &mut budget)?;
    if source_shape.headers != write.headers.len() || source_shape.footers != write.footers.len() {
        return Err(DecodeError::missing_required(
            "one replacement per header/footer reference",
        ));
    }
    let mut replacement_fields = 0usize;
    let mut replacement_bytes = 0usize;
    for replacement in write.headers.iter().chain(write.footers.iter()) {
        let fields_before = budget.fields;
        scan_reference(replacement.raw, options, &mut budget)?;
        replacement_fields = replacement_fields
            .checked_add(budget.fields.saturating_sub(fields_before))
            .ok_or_else(|| {
                DecodeError::limit(WireResourceLimit::Fields {
                    observed: usize::MAX,
                    maximum: options.max_fields,
                })
            })?;
        replacement_bytes = replacement_bytes
            .checked_add(replacement.raw.len())
            .ok_or_else(|| {
                DecodeError::limit(WireResourceLimit::WorkBytes {
                    observed: usize::MAX,
                    maximum: options.max_work_bytes,
                })
            })?;
    }

    let output_bytes = measure_template(source, write, options, &mut budget)?;
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::limit(WireResourceLimit::OutputBytes {
            observed: output_bytes,
            maximum: options.max_output_bytes,
        }));
    }

    // Check the exact candidate traversal before reserving output.  The
    // checks do not mutate the report; the candidate scan after allocation
    // charges these same counters exactly once on the shared budget.
    let candidate_fields = source_shape
        .root_fields
        .checked_add(replacement_fields)
        .ok_or_else(|| {
            DecodeError::limit(WireResourceLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            })
        })?;
    let candidate_work = output_bytes
        .saturating_mul(2)
        .checked_add(replacement_bytes.saturating_mul(2))
        .ok_or_else(|| {
            DecodeError::limit(WireResourceLimit::WorkBytes {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    budget.ensure_fields(candidate_fields)?;
    let emission_work = source.len().saturating_add(output_bytes);
    budget.ensure_work(candidate_work.saturating_add(emission_work))?;
    budget.ensure_depth(budget.max_depth.max(REFERENCE_DEPTH))?;
    budget.ensure_references(source_shape.headers.saturating_add(source_shape.footers))?;
    budget.charge_work(emission_work)?;

    let mut output = reserve_output(output_bytes)?;
    budget.record_allocation(output_bytes);
    emit_template(source, write, &mut output)?;
    if output.len() != output_bytes {
        return Err(DecodeError::projection());
    }
    let changed = source != output;
    let candidate_options = DecodeOptions {
        max_input_bytes: options.max_input_bytes.max(output.len()),
        ..options
    };
    let candidate_shape =
        scan_template_expected(&output, candidate_options, &mut budget, Some(write))?;
    if candidate_shape.headers != source_shape.headers
        || candidate_shape.footers != source_shape.footers
    {
        return Err(DecodeError::projection());
    }
    Ok((
        output,
        budget.rewrite_report(source.len(), output_bytes, changed),
    ))
}

/// Compatibility alias using the native archive name.
pub fn rewrite_section_template_archive(
    source: &[u8],
    write: SectionTemplateWrite<'_>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    rewrite_section_template(source, write, options)
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    if options.max_input_bytes > buffa::MAX_MESSAGE_BYTES as usize {
        return Err(DecodeError::limit(WireResourceLimit::InputBytes {
            observed: options.max_input_bytes,
            maximum: buffa::MAX_MESSAGE_BYTES as usize,
        }));
    }
    if source.len() > options.max_input_bytes {
        return Err(DecodeError::limit(WireResourceLimit::InputBytes {
            observed: source.len(),
            maximum: options.max_input_bytes,
        }));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::limit(WireResourceLimit::Nesting {
            observed: options.recursion_limit,
            maximum: MAX_RECURSION_LIMIT,
        }));
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct TemplateShape {
    headers: usize,
    footers: usize,
    root_fields: usize,
    reference_fields: usize,
}

fn scan_template(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<TemplateShape, DecodeError> {
    scan_template_expected(source, options, budget, None)
}

fn scan_template_expected(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    expected: Option<SectionTemplateWrite<'_>>,
) -> Result<TemplateShape, DecodeError> {
    budget.charge_message(source.len(), ROOT_DEPTH)?;
    let mut remaining = source;
    let fields_before = budget.fields;
    let mut shape = TemplateShape {
        headers: 0,
        footers: 0,
        root_fields: 0,
        reference_fields: 0,
    };
    while let Some(field) = next_field(&mut remaining, options.recursion_limit, ROOT_DEPTH, budget)?
    {
        if field.number != HEADERS_FIELD && field.number != FOOTERS_FIELD {
            continue;
        }
        let payload = field.length_delimited()?;
        if let Some(expected) = expected {
            let replacement = if field.number == HEADERS_FIELD {
                expected.headers.get(shape.headers).map(|value| value.raw)
            } else {
                expected.footers.get(shape.footers).map(|value| value.raw)
            }
            .ok_or_else(DecodeError::projection)?;
            if payload != replacement {
                return Err(DecodeError::projection());
            }
        }
        budget.charge_reference(options.max_references)?;
        let reference_fields_before = budget.fields;
        scan_reference(payload, options, budget)?;
        shape.reference_fields = shape
            .reference_fields
            .saturating_add(budget.fields.saturating_sub(reference_fields_before));
        let count = if field.number == HEADERS_FIELD {
            &mut shape.headers
        } else {
            &mut shape.footers
        };
        *count = count.checked_add(1).ok_or_else(|| {
            DecodeError::limit(WireResourceLimit::References {
                observed: usize::MAX,
                maximum: options.max_references,
            })
        })?;
    }
    shape.root_fields = budget
        .fields
        .saturating_sub(fields_before)
        .saturating_sub(shape.reference_fields);
    Ok(shape)
}

fn scan_reference<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ReferenceSnapshot<'source>, DecodeError> {
    budget.charge_message(source.len(), REFERENCE_DEPTH)?;
    let mut remaining = source;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    while let Some(field) = next_field(
        &mut remaining,
        options.recursion_limit,
        REFERENCE_DEPTH,
        budget,
    )? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
                }
                let value = field.varint()?;
                if value == 0 {
                    return Err(DecodeError::noncanonical(
                        "TSP.Reference.identifier must be nonzero",
                    ));
                }
                identifier = Some(value);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_type",
                    ));
                }
                deprecated_type = Some(require_canonical_int32(field.varint()?)?);
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if deprecated_is_external.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                deprecated_is_external = Some(require_canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    let identifier =
        identifier.ok_or_else(|| DecodeError::missing_required("TSP.Reference.identifier"))?;
    let view: projection::PagesDrawableOrderReferenceArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    if !view.has_identifier()
        || view.identifier != identifier
        || view.deprecated_type != deprecated_type
        || view.deprecated_is_external != deprecated_is_external
    {
        return Err(DecodeError::projection());
    }
    Ok(ReferenceSnapshot {
        raw: source,
        identifier,
        deprecated_type,
        deprecated_is_external,
    })
}

fn measure_template(
    source: &[u8],
    write: SectionTemplateWrite<'_>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    // Sizing is a complete second wire traversal.  Charge its source span
    // explicitly so a large interleaved unknown-field envelope cannot make
    // the pre-allocation pass appear free in the aggregate work report.
    budget.charge_work(source.len())?;
    let mut remaining = source;
    let mut header_index = 0usize;
    let mut footer_index = 0usize;
    let mut output_bytes = 0usize;
    while let Some(field) = next_field(&mut remaining, options.recursion_limit, ROOT_DEPTH, budget)?
    {
        let replacement = match field.number {
            HEADERS_FIELD => {
                let replacement = write
                    .headers
                    .get(header_index)
                    .ok_or_else(DecodeError::projection)?;
                header_index = header_index.saturating_add(1);
                replacement.raw
            },
            FOOTERS_FIELD => {
                let replacement = write
                    .footers
                    .get(footer_index)
                    .ok_or_else(DecodeError::projection)?;
                footer_index = footer_index.saturating_add(1);
                replacement.raw
            },
            _ => {
                output_bytes = output_bytes.checked_add(field.raw_len).ok_or_else(|| {
                    DecodeError::limit(WireResourceLimit::OutputBytes {
                        observed: usize::MAX,
                        maximum: options.max_output_bytes,
                    })
                })?;
                continue;
            },
        };
        let field_len = encoded_varint_len((u64::from(field.number) << 3) | 2)
            .saturating_add(encoded_varint_len(replacement.len() as u64))
            .saturating_add(replacement.len());
        output_bytes = output_bytes.checked_add(field_len).ok_or_else(|| {
            DecodeError::limit(WireResourceLimit::OutputBytes {
                observed: usize::MAX,
                maximum: options.max_output_bytes,
            })
        })?;
        if output_bytes > options.max_output_bytes {
            return Err(DecodeError::limit(WireResourceLimit::OutputBytes {
                observed: output_bytes,
                maximum: options.max_output_bytes,
            }));
        }
    }
    if header_index != write.headers.len() || footer_index != write.footers.len() {
        return Err(DecodeError::projection());
    }
    Ok(output_bytes)
}

fn emit_template(
    source: &[u8],
    write: SectionTemplateWrite<'_>,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut remaining = source;
    let mut header_index = 0usize;
    let mut footer_index = 0usize;
    while !remaining.is_empty() {
        let field = parse_unmetered_field(&mut remaining).map_err(|_| DecodeError::projection())?;
        let field = field.ok_or_else(DecodeError::projection)?;
        match field.number {
            HEADERS_FIELD => {
                let replacement = write
                    .headers
                    .get(header_index)
                    .ok_or_else(DecodeError::projection)?;
                append_length_field(output, HEADERS_FIELD, replacement.raw);
                header_index = header_index.saturating_add(1);
            },
            FOOTERS_FIELD => {
                let replacement = write
                    .footers
                    .get(footer_index)
                    .ok_or_else(DecodeError::projection)?;
                append_length_field(output, FOOTERS_FIELD, replacement.raw);
                footer_index = footer_index.saturating_add(1);
            },
            _ => output.extend_from_slice(field.raw),
        }
    }
    if header_index != write.headers.len() || footer_index != write.footers.len() {
        return Err(DecodeError::projection());
    }
    Ok(())
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
    wire_type: u8,
    value: StrictValue<'source>,
    canonical_value: bool,
    raw_len: usize,
}

impl<'source> StrictField<'source> {
    fn require_wire(self, expected: u8) -> Result<(), DecodeError> {
        if self.wire_type != expected {
            return Err(DecodeError::from(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected,
                actual: self.wire_type,
            }));
        }
        Ok(())
    }

    fn varint(self) -> Result<u64, DecodeError> {
        self.require_wire(0)?;
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("protobuf varint value"));
        }
        let StrictValue::Varint(value) = self.value else {
            return Err(DecodeError::projection());
        };
        Ok(value)
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        self.require_wire(2)?;
        let StrictValue::LengthDelimited(value) = self.value else {
            return Err(DecodeError::projection());
        };
        Ok(value)
    }
}

#[derive(Clone, Copy, Debug)]
enum ParseItem<'source> {
    Field(StrictField<'source>),
    EndGroup(u32),
}

fn next_field<'source>(
    source: &mut &'source [u8],
    recursion_limit: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_field(source, recursion_limit, depth, budget)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(number)) => Err(DecodeError::from(
            buffa::DecodeError::InvalidEndGroup(number),
        )),
        None => Ok(None),
    }
}

fn parse_field<'source>(
    source: &mut &'source [u8],
    recursion_limit: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    let before = *source;
    let (tag, canonical_key) = take_varint(source)?;
    if !canonical_key {
        return Err(DecodeError::noncanonical("protobuf field key"));
    }
    budget.charge_field()?;
    let raw_tag = u32::try_from(tag).map_err(|_| buffa::DecodeError::InvalidFieldNumber)?;
    let number = raw_tag >> 3;
    if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(DecodeError::from(buffa::DecodeError::InvalidFieldNumber));
    }
    let wire = (raw_tag & 7) as u8;
    let (value, canonical_value) = match wire {
        0 => {
            let (value, canonical) = take_varint(source)?;
            (StrictValue::Varint(value), canonical)
        },
        1 => {
            take_exact(source, 8)?;
            (StrictValue::Fixed64, true)
        },
        2 => {
            let (length, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("length-delimited size"));
            }
            let length = usize::try_from(length)
                .map_err(|_| DecodeError::from(buffa::DecodeError::MessageTooLarge))?;
            (
                StrictValue::LengthDelimited(take_exact(source, length)?),
                true,
            )
        },
        3 => {
            let child_limit = recursion_limit.checked_sub(1).ok_or_else(|| {
                DecodeError::limit(WireResourceLimit::Nesting {
                    observed: recursion_limit.saturating_add(1),
                    maximum: recursion_limit,
                })
            })?;
            skip_group(source, number, child_limit, depth.saturating_add(1), budget)?;
            (StrictValue::Group, true)
        },
        4 => return Ok(Some(ParseItem::EndGroup(number))),
        5 => {
            take_exact(source, 4)?;
            (StrictValue::Fixed32, true)
        },
        _ => {
            return Err(DecodeError::from(buffa::DecodeError::InvalidWireType(
                u32::from(wire),
            )));
        },
    };
    Ok(Some(ParseItem::Field(StrictField {
        number,
        wire_type: wire,
        value,
        canonical_value,
        raw_len: before.len() - source.len(),
    })))
}

fn skip_group(
    source: &mut &[u8],
    expected: u32,
    recursion_limit: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_depth(depth)?;
    loop {
        match parse_field(source, recursion_limit, depth, budget)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected => return Ok(()),
            Some(ParseItem::EndGroup(number)) => {
                return Err(DecodeError::from(buffa::DecodeError::InvalidEndGroup(
                    number,
                )));
            },
            None => return Err(DecodeError::from(buffa::DecodeError::UnexpectedEof)),
        }
    }
}

fn require_canonical_bool(value: u64) -> Result<bool, DecodeError> {
    if value > 1 {
        return Err(DecodeError::noncanonical("bool scalar is not zero or one"));
    }
    Ok(value == 1)
}

fn require_canonical_int32(value: u64) -> Result<i32, DecodeError> {
    if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical(
            "int32 scalar is not sign extended",
        ));
    }
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        reason = "the strict range check proves the value is a canonical int32"
    )]
    Ok(value as i32)
}

fn take_varint(source: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original
            .get(index)
            .ok_or_else(|| DecodeError::from(buffa::DecodeError::UnexpectedEof))?;
        if index == 9 && byte > 1 {
            return Err(DecodeError::from(buffa::DecodeError::VarintTooLong));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            *source = &original[consumed..];
            return Ok((value, encoded_varint_len(value) == consumed));
        }
    }
    Err(DecodeError::from(buffa::DecodeError::VarintTooLong))
}

fn take_exact<'source>(
    source: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], DecodeError> {
    if source.len() < length {
        return Err(DecodeError::from(buffa::DecodeError::UnexpectedEof));
    }
    let (selected, remaining) = source.split_at(length);
    *source = remaining;
    Ok(selected)
}

const fn encoded_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn reserve_output(amount: usize) -> Result<Vec<u8>, DecodeError> {
    #[cfg(test)]
    OUTPUT_ALLOCATIONS.with(|count| count.set(count.get().saturating_add(1)));
    let mut output = Vec::new();
    output
        .try_reserve_exact(amount)
        .map_err(|_| DecodeError::allocation(amount))?;
    if output.capacity() != amount {
        return Err(DecodeError::allocation(amount));
    }
    Ok(output)
}

#[cfg(test)]
thread_local! {
    static OUTPUT_ALLOCATIONS: Cell<usize> = const { Cell::new(0) };
}

#[cfg(test)]
fn reset_output_allocations() {
    OUTPUT_ALLOCATIONS.with(|count| count.set(0));
}

#[cfg(test)]
fn output_allocations() -> usize {
    OUTPUT_ALLOCATIONS.with(Cell::get)
}

#[derive(Clone, Copy, Debug)]
enum UnmeteredValue<'source> {
    Varint(u64),
    Fixed64,
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32,
}

#[derive(Clone, Copy, Debug)]
struct UnmeteredField<'source> {
    number: u32,
    value: UnmeteredValue<'source>,
    raw: &'source [u8],
}

fn parse_unmetered_field<'source>(
    source: &mut &'source [u8],
) -> Result<Option<UnmeteredField<'source>>, ()> {
    if source.is_empty() {
        return Ok(None);
    }
    let before = *source;
    let (tag, _) = read_varint_unmetered(source)?;
    let number = u32::try_from(tag >> 3).map_err(|_| ())?;
    if number == 0 {
        return Err(());
    }
    let wire = (tag & 7) as u8;
    let value = match wire {
        0 => UnmeteredValue::Varint(read_varint_unmetered(source)?.0),
        1 => {
            take_exact_unmetered(source, 8)?;
            UnmeteredValue::Fixed64
        },
        2 => {
            let length = usize::try_from(read_varint_unmetered(source)?.0).map_err(|_| ())?;
            UnmeteredValue::LengthDelimited(take_exact_unmetered(source, length)?)
        },
        3 => {
            skip_group_unmetered(source, number)?;
            UnmeteredValue::Group
        },
        4 => return Err(()),
        5 => {
            take_exact_unmetered(source, 4)?;
            UnmeteredValue::Fixed32
        },
        _ => return Err(()),
    };
    let raw = &before[..before.len() - source.len()];
    Ok(Some(UnmeteredField { number, value, raw }))
}

fn parse_unmetered_reference<'source>(
    source: &'source [u8],
) -> Result<ReferenceSnapshot<'source>, ()> {
    let mut remaining = source;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    while let Some(field) = parse_unmetered_field(&mut remaining)? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                let UnmeteredValue::Varint(value) = field.value else {
                    return Err(());
                };
                if identifier.replace(value).is_some() {
                    return Err(());
                }
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                let UnmeteredValue::Varint(value) = field.value else {
                    return Err(());
                };
                let value = decode_int32_unmetered(value)?;
                if deprecated_type.replace(value).is_some() {
                    return Err(());
                }
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                let UnmeteredValue::Varint(value) = field.value else {
                    return Err(());
                };
                if value > 1 || deprecated_is_external.replace(value == 1).is_some() {
                    return Err(());
                }
            },
            _ => {},
        }
    }
    let identifier = identifier.ok_or(())?;
    if identifier == 0 {
        return Err(());
    }
    Ok(ReferenceSnapshot {
        raw: source,
        identifier,
        deprecated_type,
        deprecated_is_external,
    })
}

fn read_varint_unmetered(source: &mut &[u8]) -> Result<(u64, bool), ()> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original.get(index).ok_or(())?;
        if index == 9 && byte > 1 {
            return Err(());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            *source = &original[consumed..];
            return Ok((value, encoded_varint_len(value) == consumed));
        }
    }
    Err(())
}

fn decode_int32_unmetered(value: u64) -> Result<i32, ()> {
    if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
        return Err(());
    }
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        reason = "the strict range check proves the value is a canonical int32"
    )]
    Ok(value as i32)
}

fn take_exact_unmetered<'source>(
    source: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], ()> {
    if source.len() < length {
        return Err(());
    }
    let (selected, remaining) = source.split_at(length);
    *source = remaining;
    Ok(selected)
}

fn skip_group_unmetered(source: &mut &[u8], expected: u32) -> Result<(), ()> {
    loop {
        if source.is_empty() {
            return Err(());
        }
        let (tag, _) = read_varint_unmetered(source)?;
        let number = u32::try_from(tag >> 3).map_err(|_| ())?;
        match (tag & 7) as u8 {
            0 => {
                let _ = read_varint_unmetered(source)?;
            },
            1 => {
                take_exact_unmetered(source, 8)?;
            },
            2 => {
                let length = usize::try_from(read_varint_unmetered(source)?.0).map_err(|_| ())?;
                take_exact_unmetered(source, length)?;
            },
            3 => skip_group_unmetered(source, number)?,
            4 if number == expected => return Ok(()),
            4 => return Err(()),
            5 => {
                take_exact_unmetered(source, 4)?;
            },
            _ => return Err(()),
        }
    }
}

fn append_length_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    append_varint(output, (u64::from(number) << 3) | 2);
    append_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Self {
        Self {
            input_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            references: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
            max_references: options.max_references,
            max_depth: 0,
            max_nesting: options.recursion_limit,
            allocations: 0,
            retained_bytes: source.len(),
            scratch_bytes: 0,
        }
    }

    fn charge_field(&mut self) -> Result<(), DecodeError> {
        self.charge_fields(1)
    }

    fn charge_fields(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(amount);
        if observed > self.max_fields {
            return Err(DecodeError::limit(WireResourceLimit::Fields {
                observed,
                maximum: self.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn ensure_fields(&self, amount: usize) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(amount);
        if observed > self.max_fields {
            return Err(DecodeError::limit(WireResourceLimit::Fields {
                observed,
                maximum: self.max_fields,
            }));
        }
        Ok(())
    }

    fn charge_message(&mut self, bytes: usize, depth: u32) -> Result<(), DecodeError> {
        self.charge_depth(depth)?;
        self.charge_work(bytes.saturating_mul(2))
    }

    fn charge_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.max_nesting {
            return Err(DecodeError::limit(WireResourceLimit::Nesting {
                observed: depth,
                maximum: self.max_nesting,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self.work_bytes.saturating_add(amount);
        if observed > self.max_work_bytes {
            return Err(DecodeError::limit(WireResourceLimit::WorkBytes {
                observed,
                maximum: self.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn ensure_work(&self, amount: usize) -> Result<(), DecodeError> {
        let observed = self.work_bytes.saturating_add(amount);
        if observed > self.max_work_bytes {
            return Err(DecodeError::limit(WireResourceLimit::WorkBytes {
                observed,
                maximum: self.max_work_bytes,
            }));
        }
        Ok(())
    }

    fn charge_reference(&mut self, maximum: usize) -> Result<(), DecodeError> {
        let observed = self.references.saturating_add(1);
        if observed > maximum {
            return Err(DecodeError::limit(WireResourceLimit::References {
                observed,
                maximum,
            }));
        }
        self.references = observed;
        Ok(())
    }

    fn ensure_references(&self, amount: usize) -> Result<(), DecodeError> {
        let observed = self.references.saturating_add(amount);
        if observed > self.max_references {
            return Err(DecodeError::limit(WireResourceLimit::References {
                observed,
                maximum: self.max_references,
            }));
        }
        Ok(())
    }

    fn ensure_depth(&self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.max_nesting {
            return Err(DecodeError::limit(WireResourceLimit::Nesting {
                observed: depth,
                maximum: self.max_nesting,
            }));
        }
        Ok(())
    }

    fn record_allocation(&mut self, amount: usize) {
        self.allocations = self.allocations.saturating_add(1);
        self.retained_bytes = amount;
    }

    const fn decode_report(&self) -> DecodeReport {
        DecodeReport {
            input_bytes: self.input_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
        }
    }

    const fn rewrite_report(
        &self,
        input_bytes: usize,
        output_bytes: usize,
        changed: bool,
    ) -> RewriteReport {
        RewriteReport {
            input_bytes,
            output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
            changed,
        }
    }
}

#[derive(Debug)]
struct Budget {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    references: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_references: usize,
    max_depth: u32,
    max_nesting: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::shadow_unrelated,
    reason = "focused codec tests use explicit byte fixtures"
)]
mod tests {
    use super::*;

    fn varint(mut value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return output;
            }
        }
    }

    fn field_varint(number: u32, value: u64) -> Vec<u8> {
        [varint(u64::from(number) << 3), varint(value)].concat()
    }

    fn field_bytes(number: u32, value: &[u8]) -> Vec<u8> {
        [
            varint((u64::from(number) << 3) | 2),
            varint(value.len() as u64),
            value.to_vec(),
        ]
        .concat()
    }

    fn reference(identifier: u64, tail: &[u8]) -> Vec<u8> {
        [field_varint(1, identifier), tail.to_vec()].concat()
    }

    fn template(headers: &[Vec<u8>], footers: &[Vec<u8>], tail: &[u8]) -> Vec<u8> {
        let mut source = Vec::new();
        for value in headers {
            source.extend(field_bytes(1, value));
        }
        source.extend(tail);
        for value in footers {
            source.extend(field_bytes(2, value));
        }
        source
    }

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            source.len().saturating_mul(2).max(1),
            usize::MAX,
            source.len().saturating_mul(128).max(1),
            8,
            64,
        )
    }

    #[test]
    fn borrowed_snapshot_streams_header_and_footer_records() {
        let header = reference(11, &field_varint(2, u64::MAX));
        let footer = reference(22, &field_varint(3, 1));
        let source = template(
            std::slice::from_ref(&header),
            std::slice::from_ref(&footer),
            &field_varint(99, 7),
        );
        let (snapshot, report) = decode_section_template_with_report(&source, options(&source))
            .expect("strict template");
        assert_eq!(snapshot.header_identifiers().collect::<Vec<_>>(), [11]);
        assert_eq!(snapshot.footer_identifiers().collect::<Vec<_>>(), [22]);
        assert_eq!(
            snapshot.headers().next().expect("header").raw(),
            header.as_slice()
        );
        assert_eq!(
            snapshot.footers().next().expect("footer").raw(),
            footer.as_slice()
        );
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(snapshot.header_count(), 1);
        assert_eq!(snapshot.footer_count(), 1);
        assert_eq!(report.references(), 2);
        assert_eq!(report.allocations(), 0);
    }

    #[test]
    fn unknown_root_nested_and_overlong_values_are_retained() {
        let nested_group = [0x9b, 0x03, 0x08, 0x01, 0x9c, 0x03];
        let overlong = [0x98, 0x03, 0x81, 0x00];
        let header = reference(7, &[nested_group.as_slice(), &overlong].concat());
        let root_group = [0xa3, 0x03, 0x08, 0x01, 0xa4, 0x03];
        let source = template(std::slice::from_ref(&header), &[], &root_group);
        let replacement = [ReferenceWrite::new(header.as_slice())];
        let candidate = rewrite_section_template(
            &source,
            SectionTemplateWrite::new(&replacement, &[]),
            options(&source),
        )
        .expect("rewrite");
        assert!(
            candidate
                .windows(nested_group.len())
                .any(|window| window == nested_group)
        );
        assert!(
            candidate
                .windows(overlong.len())
                .any(|window| window == overlong)
        );
        assert!(
            candidate
                .windows(root_group.len())
                .any(|window| window == root_group)
        );
    }

    #[test]
    fn malformed_reference_known_fields_fail_closed() {
        let malformed = [
            template(&[vec![0x08, 0x00]], &[], &[]),
            template(&[vec![0x10, 0x01]], &[], &[]),
            template(&[vec![0x08, 0x01, 0x08, 0x02]], &[], &[]),
            template(&[vec![0x10, 0x81, 0x00]], &[], &[]),
            template(&[vec![0x18, 0x02]], &[], &[]),
        ];
        for source in malformed {
            assert!(
                decode_section_template(&source, options(&source)).is_err(),
                "{source:?}"
            );
        }
        let wrong_root_wire = [0x08, 0x01];
        assert!(decode_section_template(&wrong_root_wire, options(&wrong_root_wire)).is_err());
    }

    #[test]
    fn rewrites_exact_lists_and_preserves_interleaved_root_bytes() {
        let first = reference(11, &[0x98, 0x03, 0x81, 0x00]);
        let second = reference(22, &[]);
        let root_tail = [field_varint(99, 7), vec![0x98, 0x03, 0x81, 0x00]].concat();
        let source = template(
            std::slice::from_ref(&first),
            std::slice::from_ref(&second),
            &root_tail,
        );
        let new_header = reference(33, &field_varint(4, 123));
        let new_footer = reference(44, &[]);
        let headers = [ReferenceWrite::new(new_header.as_slice())];
        let footers = [ReferenceWrite::new(new_footer.as_slice())];
        let candidate = rewrite_section_template(
            &source,
            SectionTemplateWrite::new(&headers, &footers),
            options(&source).with_max_output_bytes(source.len() + 32),
        )
        .expect("rewrite");
        let snapshot = decode_section_template(&candidate, options(&candidate)).expect("candidate");
        assert_eq!(snapshot.header_identifiers().collect::<Vec<_>>(), [33]);
        assert_eq!(snapshot.footer_identifiers().collect::<Vec<_>>(), [44]);
        assert!(
            candidate
                .windows(4)
                .any(|window| window == [0x98, 0x03, 0x81, 0x00])
        );
    }

    #[test]
    fn rewrite_limits_fail_before_output_allocation_and_replay_exact_report() {
        let header = reference(11, &[]);
        let footer = reference(22, &[]);
        let source = template(
            std::slice::from_ref(&header),
            std::slice::from_ref(&footer),
            &[],
        );
        let new_header = reference(111, &field_varint(2, 7));
        let new_footer = reference(222, &field_varint(2, 9));
        let headers = [ReferenceWrite::new(new_header.as_slice())];
        let footers = [ReferenceWrite::new(new_footer.as_slice())];
        let write = SectionTemplateWrite::new(&headers, &footers);
        let permissive = DecodeOptions::new(
            source.len(),
            source.len() + 64,
            usize::MAX,
            usize::MAX,
            8,
            usize::MAX,
        );
        let (expected, report) =
            rewrite_section_template_with_report(&source, write, permissive).expect("permissive");
        let exact = DecodeOptions::new(
            source.len(),
            report.output_bytes(),
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.references(),
        );
        reset_output_allocations();
        let (actual, replay) =
            rewrite_section_template_with_report(&source, write, exact).expect("exact replay");
        assert_eq!(actual, expected);
        assert_eq!(replay.output_bytes(), report.output_bytes());
        assert_eq!(replay.fields(), report.fields());
        assert_eq!(replay.work_bytes(), report.work_bytes());
        assert_eq!(output_allocations(), 1);
        for limited in [
            DecodeOptions::new(
                source.len(),
                report.output_bytes() - 1,
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.references(),
            ),
            DecodeOptions::new(
                source.len(),
                report.output_bytes(),
                report.fields() - 1,
                report.work_bytes(),
                report.max_depth(),
                report.references(),
            ),
            DecodeOptions::new(
                source.len(),
                report.output_bytes(),
                report.fields(),
                report.work_bytes() - 1,
                report.max_depth(),
                report.references(),
            ),
            DecodeOptions::new(
                source.len(),
                report.output_bytes(),
                report.fields(),
                report.work_bytes(),
                report.max_depth() - 1,
                report.references(),
            ),
            DecodeOptions::new(
                source.len(),
                report.output_bytes(),
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.references() - 1,
            ),
        ] {
            reset_output_allocations();
            assert!(rewrite_section_template_with_report(&source, write, limited).is_err());
            assert_eq!(output_allocations(), 0);
        }
    }

    #[test]
    fn unknown_group_depth_is_bounded() {
        let group = [0x9b, 0x03, 0x08, 0x01, 0x9c, 0x03];
        let source = template(&[reference(1, &[])], &[], &group);
        let exact = DecodeOptions::new(source.len(), source.len(), usize::MAX, usize::MAX, 2, 4);
        assert!(decode_section_template(&source, exact).is_ok());
        let nested = [0x9b, 0x03, 0x9b, 0x03, 0x08, 0x01, 0x9c, 0x03, 0x9c, 0x03];
        let nested_source = template(&[reference(1, &[])], &[], &nested);
        assert!(matches!(
            decode_section_template(
                &nested_source,
                DecodeOptions::new(
                    nested_source.len(),
                    nested_source.len(),
                    usize::MAX,
                    usize::MAX,
                    2,
                    4
                )
            )
            .expect_err("depth")
            .resource_limit(),
            Some(WireResourceLimit::Nesting { .. })
        ));
    }
}
