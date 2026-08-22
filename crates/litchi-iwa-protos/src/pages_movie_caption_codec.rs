//! Strict private Buffa projection for Pages movie caption metadata.
//!
//! The selected path is the bounded `TSA.CaptionInfoArchive` inheritance
//! chain used by Pages movie title/caption validation. A handwritten wire pass
//! owns framing, singularity, canonical scalar, and resource checks before a
//! private Buffa lazy view is forced. Unknown source fields remain opaque and
//! caller-owned, so this module never materializes or rewrites them.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict semantic preflight intentionally precedes the low-level wire reader."
)]

use std::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_pages_movie_caption_generated::LitchiIwaProjection as projection;

const CAPTION_INFO_SUPER_FIELD: u32 = 1;
const CAPTION_INFO_PLACEMENT_FIELD: u32 = 2;
const CAPTION_INFO_KIND_FIELD: u32 = 3;
const SHAPE_INFO_SUPER_FIELD: u32 = 1;
const SHAPE_INFO_DEPRECATED_STORAGE_FIELD: u32 = 2;
const SHAPE_INFO_OWNED_STORAGE_FIELD: u32 = 4;
const SHAPE_INFO_IS_TEXT_BOX_FIELD: u32 = 6;
const SHAPE_SUPER_FIELD: u32 = 1;
const SHAPE_STYLE_FIELD: u32 = 2;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Explicit finite resource policy for one Pages movie-caption payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    initial_recursion_limit: u32,
}

impl DecodeOptions {
    /// Build a finite bytes/fields/work/nesting policy.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_message_bytes,
            max_output_bytes: max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            initial_recursion_limit: recursion_limit,
        }
    }

    /// Replace the candidate-output ceiling used by caption-info rewrites.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the source/candidate message-byte ceiling used by readback.
    #[must_use]
    pub const fn with_max_message_bytes(mut self, maximum: usize) -> Self {
        self.max_message_bytes = maximum;
        self
    }

    /// Build a conservative policy from one known source length.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(4).max(1),
            // Caption rewrites charge source validation, sizing, emission,
            // and candidate readback against one aggregate budget. The
            // selected inheritance chain revisits nested envelopes in each
            // pass, so a 16x source cap can reject an otherwise valid
            // bounded rewrite by a few bytes.
            bytes.saturating_mul(32).max(1),
            8,
        )
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }

    fn descend(self, budget: &Budget) -> Result<Self, DecodeError> {
        let recursion_limit = self
            .recursion_limit
            .checked_sub(1)
            .ok_or_else(|| budget.recursion_limit_exceeded())?;
        Ok(Self {
            recursion_limit,
            ..self
        })
    }
}

/// Generated-free semantic facts from one caption-info payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CaptionInfoSnapshot {
    parent_identifier: u64,
    deprecated_storage_identifier: Option<u64>,
    owned_storage_identifier: Option<u64>,
    is_text_box: Option<bool>,
    style_identifier: Option<u64>,
    placement_identifier: Option<u64>,
    child_info_kind: Option<i32>,
}

impl CaptionInfoSnapshot {
    /// Required parent drawable identifier.
    #[must_use]
    pub const fn parent_identifier(self) -> u64 {
        self.parent_identifier
    }

    /// Optional legacy storage identifier.
    #[must_use]
    pub const fn deprecated_storage_identifier(self) -> Option<u64> {
        self.deprecated_storage_identifier
    }

    /// Optional owned text-storage identifier.
    #[must_use]
    pub const fn owned_storage_identifier(self) -> Option<u64> {
        self.owned_storage_identifier
    }

    /// Optional native text-box marker, preserving field presence.
    #[must_use]
    pub const fn is_text_box(self) -> Option<bool> {
        self.is_text_box
    }

    /// Optional private shape-style identifier.
    #[must_use]
    pub const fn style_identifier(self) -> Option<u64> {
        self.style_identifier
    }

    /// Optional caption-placement identifier.
    #[must_use]
    pub const fn placement_identifier(self) -> Option<u64> {
        self.placement_identifier
    }

    /// Optional native caption/title kind, including unknown enum values.
    #[must_use]
    pub const fn child_info_kind(self) -> Option<i32> {
        self.child_info_kind
    }
}

/// Failure from strict Pages movie-caption preflight or the private Buffa
/// projection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    MessageByteLimit { observed: usize, maximum: usize },
    RecursionLimit { observed: u32, maximum: u32 },
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
    OutputLimit { observed: usize, maximum: usize },
    Allocation { amount: usize },
    Projection,
}

impl DecodeError {
    /// Return the missing required schema field, when applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        let DecodeErrorKind::MissingRequired(field) = self.kind else {
            return None;
        };
        Some(field)
    }

    /// Return the duplicated singular schema field, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        let DecodeErrorKind::DuplicateSingular(field) = self.kind else {
            return None;
        };
        Some(field)
    }

    /// Return the stable canonical-wire failure reason, when applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        let DecodeErrorKind::NonCanonical(reason) = self.kind else {
            return None;
        };
        Some(reason)
    }

    /// Return the observed/configured message-byte ceiling, when applicable.
    #[must_use]
    pub const fn message_byte_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::MessageByteLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Return the observed/configured recursion ceiling, when applicable.
    #[must_use]
    pub const fn recursion_limit_values(&self) -> Option<(u32, u32)> {
        let DecodeErrorKind::RecursionLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Return the observed/configured field ceiling, when applicable.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::FieldLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Return the observed/configured work ceiling, when applicable.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::WorkLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Return the observed/configured candidate-output ceiling, when applicable.
    #[must_use]
    pub const fn output_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::OutputLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Return the requested output allocation when it could not be reserved.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        let DecodeErrorKind::Allocation { amount } = self.kind else {
            return None;
        };
        Some(amount)
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

    const fn message_byte_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::MessageByteLimit { observed, maximum },
        }
    }

    const fn recursion_limit(observed: u32, maximum: u32) -> Self {
        Self {
            kind: DecodeErrorKind::RecursionLimit { observed, maximum },
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

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    const fn output_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::OutputLimit { observed, maximum },
        }
    }

    const fn allocation(amount: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation { amount },
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::MessageByteLimit { observed, maximum } => write!(
                formatter,
                "Pages movie-caption projection byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::RecursionLimit { observed, maximum } => write!(
                formatter,
                "Pages movie-caption projection recursion limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Pages movie-caption projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Pages movie-caption projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::OutputLimit { observed, maximum } => write!(
                formatter,
                "Pages movie-caption rewrite output limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Allocation { amount } => write!(
                formatter,
                "cannot allocate Pages movie-caption rewrite output for {amount} bytes"
            ),
            DecodeErrorKind::Projection => formatter.write_str(
                "Pages movie-caption strict preflight disagrees with the Buffa projection",
            ),
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

/// Reference remapping accepted by the raw-preserving caption-info writer.
///
/// The pairs are interpreted as `old identifier -> new identifier`. Identifiers
/// absent from the slice remain unchanged. The caller owns the mapping storage;
/// no generated archive or map is materialized by the codec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CaptionInfoWrite<'mapping> {
    remap: &'mapping [(u64, u64)],
}

impl<'mapping> CaptionInfoWrite<'mapping> {
    /// Build a reference remap from borrowed identifier pairs.
    #[must_use]
    pub const fn new(remap: &'mapping [(u64, u64)]) -> Self {
        Self { remap }
    }

    fn identifier(self, current: u64) -> u64 {
        self.remap
            .iter()
            .find_map(|(old, new)| (*old == current).then_some(*new))
            .unwrap_or(current)
    }
}

/// Exact accounting for one successful caption-info rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    input_bytes: usize,
    output_bytes: usize,
    references_rewritten: usize,
}

impl RewriteReport {
    /// Number of source payload bytes inspected.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Number of candidate payload bytes produced.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Number of selected reference identifiers changed.
    #[must_use]
    pub const fn references_rewritten(self) -> usize {
        self.references_rewritten
    }
}

/// Decode the bounded Pages movie caption-info metadata projection.
pub fn decode_caption_info(
    source: &[u8],
    options: DecodeOptions,
) -> Result<CaptionInfoSnapshot, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    decode_caption_info_with_budget(source, options, &mut budget)
}

/// Decode one caption-info payload while charging an existing aggregate
/// budget. Rewrites use this entry point for both their source and candidate
/// readback passes so a work/field ceiling covers the complete transaction.
fn decode_caption_info_with_budget(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<CaptionInfoSnapshot, DecodeError> {
    let strict = preflight_caption_info(source, options, budget)?;
    let view: projection::CaptionInfoArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let projected = force_caption_info_projection(&view)?;
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(force_snapshot(strict))
}

/// Rewrite selected caption-info reference identifiers without generated
/// archive decode/encode.
///
/// Strict preflight and the private lazy projection validate the complete
/// source before any candidate bytes are allocated. The rewrite then copies
/// every unselected field span exactly, changing only canonical identifier
/// varints and the enclosing length prefixes required by those changes. A
/// strict readback validates the candidate before it is returned.
pub fn rewrite_caption_info(
    source: &[u8],
    write: CaptionInfoWrite<'_>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    Ok(rewrite_caption_info_with_report(source, write, options)?.0)
}

/// Rewrite selected caption-info references and return exact byte accounting.
pub fn rewrite_caption_info_with_report(
    source: &[u8],
    write: CaptionInfoWrite<'_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    // Decode before measuring or allocating. This keeps malformed input and
    // resource refusal failure-atomic from the caller's perspective.
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    decode_caption_info_with_budget(source, options, &mut budget)?;
    let mut changed = 0usize;
    let output_bytes = measure_caption_info(source, options, write, &mut budget, &mut changed)?;
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::output_limit(
            output_bytes,
            options.max_output_bytes,
        ));
    }
    if changed == 0 {
        let output = clone_output(source)?;
        return Ok((
            output,
            RewriteReport {
                input_bytes: source.len(),
                output_bytes: source.len(),
                references_rewritten: 0,
            },
        ));
    }

    let mut output = reserve_output(output_bytes)?;
    let mut written = 0usize;
    rewrite_caption_info_into(
        source,
        options,
        write,
        &mut budget,
        &mut output,
        &mut written,
    )?;
    debug_assert_eq!(output.len(), output_bytes);
    debug_assert_eq!(written, changed);

    let readback_options = options
        .with_max_output_bytes(options.max_output_bytes.max(output.len()))
        .with_max_message_bytes(options.max_message_bytes.max(output.len()));
    validate_decode_input(&output, readback_options)?;
    decode_caption_info_with_budget(&output, readback_options, &mut budget)?;
    Ok((
        output,
        RewriteReport {
            input_bytes: source.len(),
            output_bytes,
            references_rewritten: changed,
        },
    ))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CaptionInfoNode {
    CaptionInfo,
    ShapeInfo,
    Shape,
    Drawable,
    Reference,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RewriteTarget {
    Nested(CaptionInfoNode),
    Identifier,
}

fn selected_caption_field(
    node: CaptionInfoNode,
    number: u32,
) -> Option<(usize, RewriteTarget, &'static str, bool)> {
    match node {
        CaptionInfoNode::CaptionInfo => match number {
            CAPTION_INFO_SUPER_FIELD => Some((
                0,
                RewriteTarget::Nested(CaptionInfoNode::ShapeInfo),
                "TSA.CaptionInfoArchive.super",
                true,
            )),
            CAPTION_INFO_PLACEMENT_FIELD => Some((
                1,
                RewriteTarget::Nested(CaptionInfoNode::Reference),
                "TSA.CaptionInfoArchive.placement",
                false,
            )),
            _ => None,
        },
        CaptionInfoNode::ShapeInfo => match number {
            SHAPE_INFO_SUPER_FIELD => Some((
                0,
                RewriteTarget::Nested(CaptionInfoNode::Shape),
                "TSWP.ShapeInfoArchive.super",
                true,
            )),
            SHAPE_INFO_DEPRECATED_STORAGE_FIELD => Some((
                1,
                RewriteTarget::Nested(CaptionInfoNode::Reference),
                "TSWP.ShapeInfoArchive.deprecated_storage",
                false,
            )),
            SHAPE_INFO_OWNED_STORAGE_FIELD => Some((
                2,
                RewriteTarget::Nested(CaptionInfoNode::Reference),
                "TSWP.ShapeInfoArchive.owned_storage",
                false,
            )),
            _ => None,
        },
        CaptionInfoNode::Shape => match number {
            SHAPE_SUPER_FIELD => Some((
                0,
                RewriteTarget::Nested(CaptionInfoNode::Drawable),
                "TSD.ShapeArchive.super",
                true,
            )),
            SHAPE_STYLE_FIELD => Some((
                1,
                RewriteTarget::Nested(CaptionInfoNode::Reference),
                "TSD.ShapeArchive.style",
                false,
            )),
            _ => None,
        },
        CaptionInfoNode::Drawable => match number {
            DRAWABLE_PARENT_FIELD => Some((
                0,
                RewriteTarget::Nested(CaptionInfoNode::Reference),
                "TSD.DrawableArchive.parent",
                true,
            )),
            _ => None,
        },
        CaptionInfoNode::Reference => match number {
            REFERENCE_IDENTIFIER_FIELD => Some((
                0,
                RewriteTarget::Identifier,
                "TSP.Reference.identifier",
                true,
            )),
            _ => None,
        },
    }
}

fn measure_caption_info(
    source: &[u8],
    options: DecodeOptions,
    write: CaptionInfoWrite<'_>,
    budget: &mut Budget,
    changed: &mut usize,
) -> Result<usize, DecodeError> {
    measure_caption_node(
        source,
        options,
        write,
        budget,
        changed,
        CaptionInfoNode::CaptionInfo,
    )
}

fn measure_caption_node(
    source: &[u8],
    options: DecodeOptions,
    write: CaptionInfoWrite<'_>,
    budget: &mut Budget,
    changed: &mut usize,
    node: CaptionInfoNode,
) -> Result<usize, DecodeError> {
    budget.charge_message(source.len())?;
    let nested_options = (node != CaptionInfoNode::Reference)
        .then(|| options.descend(budget))
        .transpose()?;
    let mut seen = [false; 3];
    let mut remaining = source;
    let mut output_bytes = 0usize;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, options.recursion_limit, budget)?;
        let end = source.len() - remaining.len();
        let Some(ParseItem::Field(field)) = item else {
            return Err(match item {
                Some(ParseItem::EndGroup(number)) => {
                    buffa::DecodeError::InvalidEndGroup(number).into()
                },
                None => buffa::DecodeError::UnexpectedEof.into(),
                Some(ParseItem::Field(_)) => DecodeError::projection(),
            });
        };
        let replacement = if let Some((slot, target, name, _required)) =
            selected_caption_field(node, field.number)
        {
            if seen[slot] {
                return Err(DecodeError::duplicate_singular(name));
            }
            seen[slot] = true;
            match target {
                RewriteTarget::Nested(child) => {
                    let nested = field.length_delimited()?;
                    let child_options = nested_options.ok_or_else(DecodeError::projection)?;
                    let nested_bytes =
                        measure_caption_node(nested, child_options, write, budget, changed, child)?;
                    length_delimited_field_len(field.number, nested_bytes)
                },
                RewriteTarget::Identifier => {
                    let current = field.varint()?;
                    let identifier = write.identifier(current);
                    if identifier != current {
                        *changed = changed.saturating_add(1);
                    }
                    varint_field_len(field.number, identifier)
                },
            }
        } else {
            end - start
        };
        output_bytes = checked_output_add(output_bytes, replacement, options)?;
    }
    for &(slot, _target, name, required) in selected_caption_fields(node) {
        if required && !seen[slot] {
            return Err(DecodeError::missing_required(name));
        }
    }
    Ok(output_bytes)
}

fn selected_caption_fields(
    node: CaptionInfoNode,
) -> &'static [(usize, RewriteTarget, &'static str, bool)] {
    match node {
        CaptionInfoNode::CaptionInfo => &[
            (
                0,
                RewriteTarget::Nested(CaptionInfoNode::ShapeInfo),
                "TSA.CaptionInfoArchive.super",
                true,
            ),
            (
                1,
                RewriteTarget::Nested(CaptionInfoNode::Reference),
                "TSA.CaptionInfoArchive.placement",
                false,
            ),
        ],
        CaptionInfoNode::ShapeInfo => &[
            (
                0,
                RewriteTarget::Nested(CaptionInfoNode::Shape),
                "TSWP.ShapeInfoArchive.super",
                true,
            ),
            (
                1,
                RewriteTarget::Nested(CaptionInfoNode::Reference),
                "TSWP.ShapeInfoArchive.deprecated_storage",
                false,
            ),
            (
                2,
                RewriteTarget::Nested(CaptionInfoNode::Reference),
                "TSWP.ShapeInfoArchive.owned_storage",
                false,
            ),
        ],
        CaptionInfoNode::Shape => &[
            (
                0,
                RewriteTarget::Nested(CaptionInfoNode::Drawable),
                "TSD.ShapeArchive.super",
                true,
            ),
            (
                1,
                RewriteTarget::Nested(CaptionInfoNode::Reference),
                "TSD.ShapeArchive.style",
                false,
            ),
        ],
        CaptionInfoNode::Drawable => &[(
            0,
            RewriteTarget::Nested(CaptionInfoNode::Reference),
            "TSD.DrawableArchive.parent",
            true,
        )],
        CaptionInfoNode::Reference => &[(
            0,
            RewriteTarget::Identifier,
            "TSP.Reference.identifier",
            true,
        )],
    }
}

fn rewrite_caption_info_into(
    source: &[u8],
    options: DecodeOptions,
    write: CaptionInfoWrite<'_>,
    budget: &mut Budget,
    output: &mut Vec<u8>,
    changed: &mut usize,
) -> Result<(), DecodeError> {
    rewrite_caption_node_into(
        source,
        options,
        write,
        budget,
        output,
        changed,
        CaptionInfoNode::CaptionInfo,
    )
}

fn rewrite_caption_node_into(
    source: &[u8],
    options: DecodeOptions,
    write: CaptionInfoWrite<'_>,
    budget: &mut Budget,
    output: &mut Vec<u8>,
    changed: &mut usize,
    node: CaptionInfoNode,
) -> Result<(), DecodeError> {
    budget.charge_message(source.len())?;
    let nested_options = (node != CaptionInfoNode::Reference)
        .then(|| options.descend(budget))
        .transpose()?;
    let mut seen = [false; 3];
    let mut remaining = source;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, options.recursion_limit, budget)?;
        let end = source.len() - remaining.len();
        let Some(ParseItem::Field(field)) = item else {
            return Err(match item {
                Some(ParseItem::EndGroup(number)) => {
                    buffa::DecodeError::InvalidEndGroup(number).into()
                },
                None => buffa::DecodeError::UnexpectedEof.into(),
                Some(ParseItem::Field(_)) => DecodeError::projection(),
            });
        };
        let Some((slot, target, name, _required)) = selected_caption_field(node, field.number)
        else {
            output.extend_from_slice(&source[start..end]);
            continue;
        };
        if seen[slot] {
            return Err(DecodeError::duplicate_singular(name));
        }
        seen[slot] = true;
        match target {
            RewriteTarget::Nested(child) => {
                let nested = field.length_delimited()?;
                let child_options = nested_options.ok_or_else(DecodeError::projection)?;
                let mut nested_changed = 0usize;
                let nested_bytes = measure_caption_node(
                    nested,
                    child_options,
                    write,
                    budget,
                    &mut nested_changed,
                    child,
                )?;
                let mut nested_output = reserve_output(nested_bytes)?;
                rewrite_caption_node_into(
                    nested,
                    child_options,
                    write,
                    budget,
                    &mut nested_output,
                    changed,
                    child,
                )?;
                debug_assert_eq!(nested_output.len(), nested_bytes);
                append_length_delimited_field(output, field.number, &nested_output);
            },
            RewriteTarget::Identifier => {
                let current = field.varint()?;
                let identifier = write.identifier(current);
                if identifier != current {
                    *changed = changed.saturating_add(1);
                }
                append_varint_field(output, field.number, identifier);
            },
        }
    }
    for &(slot, _target, name, required) in selected_caption_fields(node) {
        if required && !seen[slot] {
            return Err(DecodeError::missing_required(name));
        }
    }
    Ok(())
}

fn checked_output_add(
    current: usize,
    additional: usize,
    options: DecodeOptions,
) -> Result<usize, DecodeError> {
    let total = current
        .checked_add(additional)
        .ok_or_else(|| DecodeError::output_limit(usize::MAX, options.max_output_bytes))?;
    if total > options.max_output_bytes {
        return Err(DecodeError::output_limit(total, options.max_output_bytes));
    }
    Ok(total)
}

fn length_delimited_field_len(number: u32, payload_len: usize) -> usize {
    varint_len((u64::from(number) << 3) | 2) + varint_len(payload_len as u64) + payload_len
}

fn varint_field_len(number: u32, value: u64) -> usize {
    varint_len(u64::from(number) << 3) + varint_len(value)
}

fn reserve_output(amount: usize) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(amount)
        .map_err(|_allocation_error| DecodeError::allocation(amount))?;
    if output.capacity() != amount {
        return Err(DecodeError::allocation(amount));
    }
    Ok(output)
}

fn clone_output(source: &[u8]) -> Result<Vec<u8>, DecodeError> {
    let mut output = reserve_output(source.len())?;
    output.extend_from_slice(source);
    Ok(output)
}

fn append_length_delimited_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    append_varint(output, (u64::from(number) << 3) | 2);
    append_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    append_varint(output, u64::from(number) << 3);
    append_varint(output, value);
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let max_buffa_message_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| DecodeError::message_byte_limit(options.max_message_bytes, 0))?;
    if options.max_message_bytes > max_buffa_message_bytes {
        return Err(DecodeError::message_byte_limit(
            options.max_message_bytes,
            max_buffa_message_bytes,
        ));
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::message_byte_limit(
            source.len(),
            options.max_message_bytes,
        ));
    }
    if options.recursion_limit == 0 {
        return Err(DecodeError::recursion_limit(1, 0));
    }
    if options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::recursion_limit(
            options.recursion_limit,
            MAX_RECURSION_LIMIT,
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
    max_recursion_limit: u32,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
            max_recursion_limit: options.initial_recursion_limit,
        }
    }

    fn charge_field(&mut self) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(1);
        if observed > self.max_fields {
            return Err(DecodeError::field_limit(observed, self.max_fields));
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_message(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self.work_bytes.saturating_add(bytes.saturating_mul(2));
        if observed > self.max_work_bytes {
            return Err(DecodeError::work_limit(observed, self.max_work_bytes));
        }
        self.work_bytes = observed;
        Ok(())
    }

    const fn recursion_limit_exceeded(&self) -> DecodeError {
        DecodeError::recursion_limit(
            self.max_recursion_limit.saturating_add(1),
            self.max_recursion_limit,
        )
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ReferenceSnapshot {
    identifier: u64,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct StrictCaptionInfo {
    parent: ReferenceSnapshot,
    deprecated_storage: Option<ReferenceSnapshot>,
    owned_storage: Option<ReferenceSnapshot>,
    is_text_box: Option<bool>,
    style: Option<ReferenceSnapshot>,
    placement: Option<ReferenceSnapshot>,
    child_info_kind: Option<i32>,
}

fn preflight_caption_info(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<StrictCaptionInfo, DecodeError> {
    budget.charge_message(source.len())?;
    let nested = options.descend(budget)?;
    let mut super_info = None;
    let mut placement = None;
    let mut child_info_kind = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        match field.number {
            CAPTION_INFO_SUPER_FIELD => {
                if super_info.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSA.CaptionInfoArchive.super",
                    ));
                }
                super_info = Some(preflight_shape_info(
                    field.length_delimited()?,
                    nested,
                    budget,
                )?);
            },
            CAPTION_INFO_PLACEMENT_FIELD => {
                if placement.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSA.CaptionInfoArchive.placement",
                    ));
                }
                placement = Some(preflight_reference(
                    field.length_delimited()?,
                    nested,
                    budget,
                )?);
            },
            CAPTION_INFO_KIND_FIELD => {
                if child_info_kind.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSA.CaptionInfoArchive.childInfoKind",
                    ));
                }
                child_info_kind = Some(decode_int32(require_canonical_int32(field.varint()?)?));
            },
            _ => {},
        }
    }
    let super_info =
        super_info.ok_or_else(|| DecodeError::missing_required("TSA.CaptionInfoArchive.super"))?;
    Ok(StrictCaptionInfo {
        parent: super_info.parent,
        deprecated_storage: super_info.deprecated_storage,
        owned_storage: super_info.owned_storage,
        is_text_box: super_info.is_text_box,
        style: super_info.style,
        placement,
        child_info_kind,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct StrictShapeInfo {
    parent: ReferenceSnapshot,
    deprecated_storage: Option<ReferenceSnapshot>,
    owned_storage: Option<ReferenceSnapshot>,
    is_text_box: Option<bool>,
    style: Option<ReferenceSnapshot>,
}

fn preflight_shape_info(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<StrictShapeInfo, DecodeError> {
    budget.charge_message(source.len())?;
    let nested = options.descend(budget)?;
    let mut shape = None;
    let mut deprecated_storage = None;
    let mut owned_storage = None;
    let mut is_text_box = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        match field.number {
            SHAPE_INFO_SUPER_FIELD => {
                if shape.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.ShapeInfoArchive.super",
                    ));
                }
                shape = Some(preflight_shape(field.length_delimited()?, nested, budget)?);
            },
            SHAPE_INFO_DEPRECATED_STORAGE_FIELD => {
                if deprecated_storage.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.ShapeInfoArchive.deprecated_storage",
                    ));
                }
                deprecated_storage = Some(preflight_reference(
                    field.length_delimited()?,
                    nested,
                    budget,
                )?);
            },
            SHAPE_INFO_OWNED_STORAGE_FIELD => {
                if owned_storage.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.ShapeInfoArchive.owned_storage",
                    ));
                }
                owned_storage = Some(preflight_reference(
                    field.length_delimited()?,
                    nested,
                    budget,
                )?);
            },
            SHAPE_INFO_IS_TEXT_BOX_FIELD => {
                if is_text_box.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.ShapeInfoArchive.is_text_box",
                    ));
                }
                is_text_box = Some(require_canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    let shape =
        shape.ok_or_else(|| DecodeError::missing_required("TSWP.ShapeInfoArchive.super"))?;
    Ok(StrictShapeInfo {
        parent: shape.parent,
        deprecated_storage,
        owned_storage,
        is_text_box,
        style: shape.style,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct StrictShape {
    parent: ReferenceSnapshot,
    style: Option<ReferenceSnapshot>,
}

fn preflight_shape(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<StrictShape, DecodeError> {
    budget.charge_message(source.len())?;
    let nested = options.descend(budget)?;
    let mut drawable = None;
    let mut style = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        match field.number {
            SHAPE_SUPER_FIELD => {
                if drawable.is_some() {
                    return Err(DecodeError::duplicate_singular("TSD.ShapeArchive.super"));
                }
                drawable = Some(preflight_drawable(
                    field.length_delimited()?,
                    nested,
                    budget,
                )?);
            },
            SHAPE_STYLE_FIELD => {
                if style.is_some() {
                    return Err(DecodeError::duplicate_singular("TSD.ShapeArchive.style"));
                }
                style = Some(preflight_reference(
                    field.length_delimited()?,
                    nested,
                    budget,
                )?);
            },
            _ => {},
        }
    }
    let parent = drawable.ok_or_else(|| DecodeError::missing_required("TSD.ShapeArchive.super"))?;
    Ok(StrictShape { parent, style })
}

fn preflight_drawable(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ReferenceSnapshot, DecodeError> {
    budget.charge_message(source.len())?;
    let nested = options.descend(budget)?;
    let mut parent = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        if field.number != DRAWABLE_PARENT_FIELD {
            continue;
        }
        if parent.is_some() {
            return Err(DecodeError::duplicate_singular(
                "TSD.DrawableArchive.parent",
            ));
        }
        parent = Some(preflight_reference(
            field.length_delimited()?,
            nested,
            budget,
        )?);
    }
    parent.ok_or_else(|| DecodeError::missing_required("TSD.DrawableArchive.parent"))
}

fn preflight_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ReferenceSnapshot, DecodeError> {
    budget.charge_message(source.len())?;
    let mut identifier = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        if field.number != REFERENCE_IDENTIFIER_FIELD {
            continue;
        }
        if identifier.is_some() {
            return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
        }
        identifier = Some(field.varint()?);
    }
    Ok(ReferenceSnapshot {
        identifier: identifier
            .ok_or_else(|| DecodeError::missing_required("TSP.Reference.identifier"))?,
    })
}

fn force_caption_info_projection(
    view: &projection::CaptionInfoArchiveLazyView<'_>,
) -> Result<StrictCaptionInfo, DecodeError> {
    let shape_info = view
        .super_
        .get()
        .map_err(DecodeError::from)?
        .ok_or_else(|| DecodeError::missing_required("TSA.CaptionInfoArchive.super"))?;
    let shape = shape_info
        .super_
        .get()
        .map_err(DecodeError::from)?
        .ok_or_else(|| DecodeError::missing_required("TSWP.ShapeInfoArchive.super"))?;
    let drawable = shape
        .super_
        .get()
        .map_err(DecodeError::from)?
        .ok_or_else(|| DecodeError::missing_required("TSD.ShapeArchive.super"))?;
    let parent = drawable
        .parent
        .get()
        .map_err(DecodeError::from)?
        .map(|reference| force_reference_projection(&reference))
        .transpose()?
        .ok_or_else(|| DecodeError::missing_required("TSD.DrawableArchive.parent"))?;
    let deprecated_storage = shape_info
        .deprecated_storage
        .get()
        .map_err(DecodeError::from)?
        .map(|reference| force_reference_projection(&reference))
        .transpose()?;
    let owned_storage = shape_info
        .owned_storage
        .get()
        .map_err(DecodeError::from)?
        .map(|reference| force_reference_projection(&reference))
        .transpose()?;
    let style = shape
        .style
        .get()
        .map_err(DecodeError::from)?
        .map(|reference| force_reference_projection(&reference))
        .transpose()?;
    let placement = view
        .placement
        .get()
        .map_err(DecodeError::from)?
        .map(|reference| force_reference_projection(&reference))
        .transpose()?;
    Ok(StrictCaptionInfo {
        parent,
        deprecated_storage,
        owned_storage,
        is_text_box: shape_info.is_text_box,
        style,
        placement,
        child_info_kind: view.child_info_kind,
    })
}

fn force_reference_projection(
    view: &projection::ReferenceLazyView<'_>,
) -> Result<ReferenceSnapshot, DecodeError> {
    if !view.has_identifier() {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    Ok(ReferenceSnapshot {
        identifier: view.identifier,
    })
}

fn require_canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}

fn require_canonical_int32(value: u64) -> Result<u64, DecodeError> {
    if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical(
            "int32 scalar is not a sign-extended 32-bit value",
        ));
    }
    Ok(value)
}

#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    reason = "Strict preflight proved the u64 is a canonical sign-extended int32."
)]
const fn decode_int32(value: u64) -> i32 {
    value as i32
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
    canonical_key: bool,
    canonical_value: bool,
}

impl<'source> StrictField<'source> {
    fn require_wire_type(self, expected: buffa::encoding::WireType) -> Result<(), DecodeError> {
        if !self.canonical_key {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
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
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("protobuf varint value"));
        }
        let StrictValue::Varint(value) = self.value else {
            return Err(DecodeError::projection());
        };
        Ok(value)
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::LengthDelimited)?;
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("length-delimited size"));
        }
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

fn next_strict_field<'source>(
    source: &mut &'source [u8],
    recursion_limit: u32,
    budget: &mut Budget,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_strict_field(source, recursion_limit, budget)? {
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
    let (value, canonical_value) = match wire_type {
        buffa::encoding::WireType::Varint => {
            let (value, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            }
            (StrictValue::Varint(value), canonical)
        },
        buffa::encoding::WireType::Fixed64 => {
            take_exact(source, 8)?;
            (StrictValue::Fixed64, true)
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (encoded_length, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("length-delimited size"));
            }
            let length = usize::try_from(encoded_length)
                .map_err(|_error| buffa::DecodeError::MessageTooLarge)?;
            (
                StrictValue::LengthDelimited(take_exact(source, length)?),
                canonical,
            )
        },
        buffa::encoding::WireType::StartGroup => {
            let child_limit = recursion_limit.checked_sub(1).ok_or_else(|| {
                DecodeError::recursion_limit(recursion_limit + 1, recursion_limit)
            })?;
            skip_strict_group(source, field_number, child_limit, budget)?;
            (StrictValue::Group, true)
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            take_exact(source, 4)?;
            (StrictValue::Fixed32, true)
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw_wire_type).into()),
    };
    Ok(Some(ParseItem::Field(StrictField {
        number: field_number,
        wire_type,
        value,
        canonical_key,
        canonical_value,
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

fn force_snapshot(strict: StrictCaptionInfo) -> CaptionInfoSnapshot {
    CaptionInfoSnapshot {
        parent_identifier: strict.parent.identifier,
        deprecated_storage_identifier: strict
            .deprecated_storage
            .map(|reference| reference.identifier),
        owned_storage_identifier: strict.owned_storage.map(|reference| reference.identifier),
        is_text_box: strict.is_text_box,
        style_identifier: strict.style.map(|reference| reference.identifier),
        placement_identifier: strict.placement.map(|reference| reference.identifier),
        child_info_kind: strict.child_info_kind,
    }
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::same_item_push,
    reason = "Focused wire fixtures intentionally use direct assertions and byte builders."
)]
mod tests {
    use super::{
        Budget, CAPTION_INFO_KIND_FIELD, CaptionInfoWrite, DecodeOptions, decode_caption_info,
        decode_caption_info_with_budget, measure_caption_info, reserve_output,
        rewrite_caption_info_into, rewrite_caption_info_with_report, validate_decode_input,
    };

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

    fn reference(identifier: u64) -> Vec<u8> {
        [field_varint(1, identifier), field_varint(90, 7)].concat()
    }

    fn caption_info() -> Vec<u8> {
        caption_info_with_text_box(1)
    }

    fn caption_info_with_text_box(value: u64) -> Vec<u8> {
        let drawable = field_bytes(2, &reference(11));
        let shape = [field_bytes(1, &drawable), field_bytes(2, &reference(22))].concat();
        let shape_info = [
            field_bytes(1, &shape),
            field_bytes(2, &reference(33)),
            field_bytes(4, &reference(44)),
            field_varint(6, value),
        ]
        .concat();
        [
            field_bytes(1, &shape_info),
            field_bytes(2, &reference(55)),
            field_varint(CAPTION_INFO_KIND_FIELD, 1),
        ]
        .concat()
    }

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
    }

    fn aggregate_rewrite_work(source: &[u8], write: CaptionInfoWrite<'_>) -> (usize, usize) {
        let options = DecodeOptions::new(source.len(), usize::MAX, usize::MAX, 8)
            .with_max_output_bytes(source.len() * 8);
        let mut budget = Budget::new(options);
        decode_caption_info_with_budget(source, options, &mut budget).expect("source decode");
        let mut changed = 0;
        let output_bytes = measure_caption_info(source, options, write, &mut budget, &mut changed)
            .expect("rewrite measure");
        let mut output = reserve_output(output_bytes).expect("rewrite allocation");
        let mut written = 0;
        rewrite_caption_info_into(
            source,
            options,
            write,
            &mut budget,
            &mut output,
            &mut written,
        )
        .expect("rewrite pass");
        let readback_options = options
            .with_max_message_bytes(options.max_message_bytes.max(output.len()))
            .with_max_output_bytes(options.max_output_bytes.max(output.len()));
        validate_decode_input(&output, readback_options).expect("readback input");
        decode_caption_info_with_budget(&output, readback_options, &mut budget)
            .expect("rewrite readback");
        assert_eq!(written, changed);
        (output_bytes, budget.work_bytes)
    }

    #[test]
    fn selected_caption_info_chain_matches_private_projection() {
        let source = caption_info();
        let snapshot = decode_caption_info(&source, options(&source)).expect("caption info");
        assert_eq!(snapshot.parent_identifier(), 11);
        assert_eq!(snapshot.deprecated_storage_identifier(), Some(33));
        assert_eq!(snapshot.owned_storage_identifier(), Some(44));
        assert_eq!(snapshot.is_text_box(), Some(true));
        assert_eq!(snapshot.style_identifier(), Some(22));
        assert_eq!(snapshot.placement_identifier(), Some(55));
        assert_eq!(snapshot.child_info_kind(), Some(1));
    }

    #[test]
    fn unknown_fields_are_wire_checked_but_not_materialized() {
        let mut source = caption_info();
        source.extend(field_varint(1000, 9));
        let before = source.clone();
        assert_eq!(
            decode_caption_info(&source, options(&source))
                .expect("unknown caption info")
                .parent_identifier(),
            11
        );
        assert_eq!(source, before);
    }

    #[test]
    fn required_chain_duplicates_and_wrong_wire_are_rejected() {
        assert_eq!(
            decode_caption_info(&[], options(&[]))
                .expect_err("missing super")
                .missing_required_field(),
            Some("TSA.CaptionInfoArchive.super")
        );
        let mut duplicate = caption_info();
        duplicate.extend(field_varint(CAPTION_INFO_KIND_FIELD, 2));
        assert_eq!(
            decode_caption_info(&duplicate, options(&duplicate))
                .expect_err("duplicate kind")
                .duplicate_singular_field(),
            Some("TSA.CaptionInfoArchive.childInfoKind")
        );
        let wrong_wire = field_bytes(CAPTION_INFO_KIND_FIELD, &[1]);
        assert!(decode_caption_info(&wrong_wire, options(&wrong_wire)).is_err());
    }

    #[test]
    fn canonical_scalars_and_limits_fail_before_buffa() {
        let source = caption_info();
        let too_large = DecodeOptions::new(source.len() - 1, source.len(), source.len() * 16, 8);
        assert_eq!(
            decode_caption_info(&source, too_large)
                .expect_err("byte limit")
                .message_byte_limit_values(),
            Some((source.len(), source.len() - 1))
        );
        let mut noncanonical = source.clone();
        noncanonical.extend([0x80, 0x00]);
        assert_eq!(
            decode_caption_info(&noncanonical, options(&noncanonical))
                .expect_err("malformed unknown key")
                .noncanonical_reason(),
            Some("protobuf field key")
        );
        let bad_bool = caption_info_with_text_box(2);
        assert_eq!(
            decode_caption_info(&bad_bool, options(&bad_bool))
                .expect_err("noncanonical bool")
                .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
    }

    #[test]
    fn rewrite_remaps_selected_references_and_preserves_unknown_spans() {
        let mut source = caption_info();
        source.extend(field_varint(1000, 9));
        let before = source.clone();
        let remap = [
            (11, 1_u64 << 40),
            (22, 1_u64 << 41),
            (33, 1_u64 << 42),
            (44, 1_u64 << 43),
            (55, 1_u64 << 44),
        ];
        let rewrite_options = options(&source).with_max_output_bytes(source.len() * 8);
        let (rewritten, report) = rewrite_caption_info_with_report(
            &source,
            CaptionInfoWrite::new(&remap),
            rewrite_options,
        )
        .expect("caption-info rewrite");
        assert_eq!(report.input_bytes(), source.len());
        assert_eq!(report.output_bytes(), rewritten.len());
        assert_eq!(report.references_rewritten(), 5);
        assert_eq!(source, before);
        let snapshot = decode_caption_info(&rewritten, options(&rewritten)).expect("readback");
        assert_eq!(snapshot.parent_identifier(), 1_u64 << 40);
        assert_eq!(snapshot.style_identifier(), Some(1_u64 << 41));
        assert_eq!(snapshot.deprecated_storage_identifier(), Some(1_u64 << 42));
        assert_eq!(snapshot.owned_storage_identifier(), Some(1_u64 << 43));
        assert_eq!(snapshot.placement_identifier(), Some(1_u64 << 44));
        assert!(
            rewritten
                .windows(field_varint(90, 7).len())
                .any(|window| window == field_varint(90, 7).as_slice())
        );
        assert!(
            rewritten
                .windows(field_varint(1000, 9).len())
                .any(|window| window == field_varint(1000, 9).as_slice())
        );
    }

    #[test]
    fn matching_rewrite_is_exact_noop_and_output_limits_refuse_before_write() {
        let source = caption_info();
        let (rewritten, report) = rewrite_caption_info_with_report(
            &source,
            CaptionInfoWrite::new(&[(11, 11)]),
            options(&source),
        )
        .expect("no-op rewrite");
        assert_eq!(rewritten, source);
        assert_eq!(report.references_rewritten(), 0);

        let error = rewrite_caption_info_with_report(
            &source,
            CaptionInfoWrite::new(&[(11, 1_u64 << 40)]),
            options(&source),
        )
        .expect_err("output limit");
        assert!(error.output_limit_values().is_some());
        assert_eq!(source, caption_info());
    }

    #[test]
    fn output_reservation_enforces_exact_capacity_or_typed_allocation_error() {
        match reserve_output(17) {
            Ok(output) => assert_eq!(output.capacity(), 17),
            Err(error) => assert_eq!(error.allocation_amount(), Some(17)),
        }
    }

    #[test]
    fn rewrite_work_budget_is_aggregate_and_exactly_inclusive() {
        let source = caption_info();
        let write = CaptionInfoWrite::new(&[
            (11, 1_u64 << 40),
            (22, 1_u64 << 41),
            (33, 1_u64 << 42),
            (44, 1_u64 << 43),
            (55, 1_u64 << 44),
        ]);
        let (output_bytes, exact_work) = aggregate_rewrite_work(&source, write);

        let exact_options = DecodeOptions::new(source.len(), usize::MAX, exact_work, 8)
            .with_max_output_bytes(output_bytes);
        rewrite_caption_info_with_report(&source, write, exact_options)
            .expect("the exact aggregate work ceiling is inclusive");

        let below_options = DecodeOptions::new(source.len(), usize::MAX, exact_work - 1, 8)
            .with_max_output_bytes(output_bytes);
        let error = rewrite_caption_info_with_report(&source, write, below_options)
            .expect_err("one byte below aggregate work must fail");
        assert_eq!(
            error.work_limit_values(),
            Some((exact_work, exact_work - 1))
        );
    }
}
