//! Strict private Buffa projection for the Pages footnote-reference edge.
//!
//! The handwritten pass validates the selected known fields and every source
//! field's wire framing before Buffa observes the payload.  Buffa then
//! cross-checks the borrowed lazy view.  Unknown fields and all source bytes
//! remain owned by the caller; this codec has no production encoding path.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict semantic preflight intentionally precedes the low-level wire reader."
)]

use std::{fmt, num::NonZeroU64, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_pages_footnote_generated::LitchiIwaProjection as projection;

const FOOTNOTE_SUPER_FIELD: u32 = 1;
const FOOTNOTE_CONTAINED_STORAGE_FIELD: u32 = 2;
const FOOTNOTE_CUSTOM_MARK_FIELD: u32 = 3;
const TEXTUAL_STRING_EQUIVALENT_FIELD: u32 = 1;
const TEXTUAL_KIND_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Explicit finite resource policy for one focused Pages footnote payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
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
            max_fields,
            max_work_bytes,
            recursion_limit,
        }
    }

    /// Build a conservative finite policy from one already-borrowed source.
    ///
    /// Callers that have a wider aggregate budget should use [`Self::new`]
    /// instead. This helper is intentionally source-sized so a direct codec
    /// caller cannot accidentally hand the lazy Buffa projection an
    /// unbounded message ceiling.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(4).max(1),
            bytes.saturating_mul(16).max(1),
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

    fn descend(self) -> Result<Self, DecodeError> {
        let maximum = self.recursion_limit;
        let recursion_limit = self
            .recursion_limit
            .checked_sub(1)
            .ok_or_else(|| DecodeError::recursion_limit(maximum.saturating_add(1), maximum))?;
        Ok(Self {
            recursion_limit,
            ..self
        })
    }
}

/// Exact generated-free projection of one `TSP.Reference`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReferenceSnapshot {
    identifier: NonZeroU64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

impl ReferenceSnapshot {
    /// Native object identifier, proven non-zero by strict preflight.
    #[must_use]
    pub const fn identifier(self) -> NonZeroU64 {
        self.identifier
    }

    /// Deprecated native object-type hint.
    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }

    /// Deprecated external-reference marker.
    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }
}

/// Borrowed, generated-free facts from one Pages footnote reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FootnoteReferenceSnapshot<'source> {
    super_string_equivalent: Option<&'source str>,
    super_kind: Option<i32>,
    contained_storage: Option<ReferenceSnapshot>,
    custom_mark_string: Option<&'source str>,
}

impl<'source> FootnoteReferenceSnapshot<'source> {
    /// Optional textual attachment string-equivalent.
    #[must_use]
    pub const fn super_string_equivalent(self) -> Option<&'source str> {
        self.super_string_equivalent
    }

    /// Optional native textual attachment kind.
    #[must_use]
    pub const fn super_kind(self) -> Option<i32> {
        self.super_kind
    }

    /// Optional contained-storage reference.
    #[must_use]
    pub const fn contained_storage(self) -> Option<ReferenceSnapshot> {
        self.contained_storage
    }

    /// Optional custom marker borrowed from the source payload.
    #[must_use]
    pub const fn custom_mark_string(self) -> Option<&'source str> {
        self.custom_mark_string
    }
}

/// Failure from strict Pages footnote preflight or its Buffa cross-check.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

/// Finite resource rejected by the Pages footnote codec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Source or configured Buffa message bytes exceeded the ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Strict wire-field visits exceeded the ceiling.
    Fields { observed: usize, maximum: usize },
    /// Strict plus lazy-projection work exceeded the ceiling.
    Work { observed: usize, maximum: usize },
    /// Configured or traversed protobuf nesting exceeded the ceiling.
    Nesting { observed: u32, maximum: u32 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Resource(DecodeLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    ZeroIdentifier(&'static str),
    InvalidUtf8(&'static str),
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
    Projection,
}

impl DecodeError {
    const fn recursion_limit(observed: u32, maximum: u32) -> Self {
        Self {
            kind: DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }),
        }
    }

    const fn resource(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Resource(limit),
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

    const fn zero_identifier(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::ZeroIdentifier(field),
        }
    }

    const fn invalid_utf8(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::InvalidUtf8(field),
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

    /// Required field absent from the source, when applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::InvalidUtf8(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Singular known field repeated in the source, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::InvalidUtf8(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Stable canonicality failure reason, when applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::InvalidUtf8(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Reference field carrying a forbidden zero identifier, when applicable.
    #[must_use]
    pub const fn zero_identifier_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::ZeroIdentifier(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::InvalidUtf8(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Text field carrying invalid UTF-8, when applicable.
    #[must_use]
    pub const fn invalid_utf8_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::InvalidUtf8(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Observed/configured field counts for a field-limit failure.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::FieldLimit { observed, maximum } => Some((observed, maximum)),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::InvalidUtf8(_)
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Observed/configured work bytes for a work-limit failure.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::WorkLimit { observed, maximum } => Some((observed, maximum)),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::InvalidUtf8(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Return the typed finite resource failure, when applicable.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Resource(limit) => Some(limit),
            DecodeErrorKind::FieldLimit { observed, maximum } => {
                Some(DecodeLimit::Fields { observed, maximum })
            },
            DecodeErrorKind::WorkLimit { observed, maximum } => {
                Some(DecodeLimit::Work { observed, maximum })
            },
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::InvalidUtf8(_)
            | DecodeErrorKind::Projection => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Resource(DecodeLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "Pages footnote projection byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "Pages footnote projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "Pages footnote projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Pages footnote projection nesting limit exceeded: observed {observed}, maximum {maximum}"
            ),
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
            DecodeErrorKind::InvalidUtf8(field) => write!(formatter, "{field} is invalid UTF-8"),
            DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Pages footnote projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Pages footnote projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Projection => formatter
                .write_str("Pages footnote strict preflight disagrees with the Buffa projection"),
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

/// Decode one complete `TSWP.FootnoteReferenceAttachmentArchive`.
///
/// The optional `super` and `contained_storage` edges retain their native
/// presence semantics.  A present contained-storage reference is required to
/// contain a non-zero identifier, matching the identity checks performed by
/// the Pages package adapter.
pub fn decode_footnote_reference<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<FootnoteReferenceSnapshot<'source>, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_footnote(source, options, &mut budget)?;
    let view: projection::FootnoteReferenceAttachmentArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let projected = force_footnote_projection(&view)?;
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let max_buffa_message_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
    if options.max_message_bytes > max_buffa_message_bytes {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: options.max_message_bytes,
            maximum: max_buffa_message_bytes,
        }));
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: options.max_message_bytes,
        }));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
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
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
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
        let strict_and_projection = bytes.saturating_mul(2);
        let observed = self.work_bytes.saturating_add(strict_and_projection);
        if observed > self.max_work_bytes {
            return Err(DecodeError::work_limit(observed, self.max_work_bytes));
        }
        self.work_bytes = observed;
        Ok(())
    }
}

fn preflight_footnote<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<FootnoteReferenceSnapshot<'source>, DecodeError> {
    budget.charge_message(source.len())?;
    let mut super_payload = None;
    let mut contained_storage = None;
    let mut custom_mark_string = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            FOOTNOTE_SUPER_FIELD => {
                if super_payload.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.FootnoteReferenceAttachmentArchive.super",
                    ));
                }
                let nested_options = options.descend()?;
                super_payload = Some(preflight_textual(
                    field.length_delimited()?,
                    nested_options,
                    budget,
                )?);
            },
            FOOTNOTE_CONTAINED_STORAGE_FIELD => {
                if contained_storage.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.FootnoteReferenceAttachmentArchive.contained_storage",
                    ));
                }
                let nested_options = options.descend()?;
                contained_storage = Some(preflight_reference(
                    field.length_delimited()?,
                    nested_options,
                    budget,
                )?);
            },
            FOOTNOTE_CUSTOM_MARK_FIELD => {
                if custom_mark_string.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.FootnoteReferenceAttachmentArchive.custom_mark_string",
                    ));
                }
                let bytes = field.length_delimited()?;
                custom_mark_string = Some(str::from_utf8(bytes).map_err(|_error| {
                    DecodeError::invalid_utf8(
                        "TSWP.FootnoteReferenceAttachmentArchive.custom_mark_string",
                    )
                })?);
            },
            _ => {},
        }
    }
    Ok(FootnoteReferenceSnapshot {
        super_string_equivalent: super_payload.and_then(|textual| textual.string_equivalent),
        super_kind: super_payload.and_then(|textual| textual.kind),
        contained_storage,
        custom_mark_string,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct TextualSnapshot<'source> {
    string_equivalent: Option<&'source str>,
    kind: Option<i32>,
}

fn preflight_textual<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<TextualSnapshot<'source>, DecodeError> {
    budget.charge_message(source.len())?;
    let mut string_equivalent = None;
    let mut kind = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            TEXTUAL_STRING_EQUIVALENT_FIELD => {
                if string_equivalent.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.TextualAttachmentArchive.string_equivalent",
                    ));
                }
                string_equivalent =
                    Some(str::from_utf8(field.length_delimited()?).map_err(|_error| {
                        DecodeError::invalid_utf8("TSWP.TextualAttachmentArchive.string_equivalent")
                    })?);
            },
            TEXTUAL_KIND_FIELD => {
                if kind.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.TextualAttachmentArchive.kind",
                    ));
                }
                kind = Some(decode_int32(require_canonical_int32(field.varint()?)?));
            },
            _ => {},
        }
    }
    Ok(TextualSnapshot {
        string_equivalent,
        kind,
    })
}

fn preflight_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ReferenceSnapshot, DecodeError> {
    budget.charge_message(source.len())?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
                }
                identifier = Some(
                    NonZeroU64::new(field.varint()?)
                        .ok_or_else(|| DecodeError::zero_identifier("TSP.Reference.identifier"))?,
                );
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_type",
                    ));
                }
                deprecated_type = Some(decode_int32(require_canonical_int32(field.varint()?)?));
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
    Ok(ReferenceSnapshot {
        identifier: identifier
            .ok_or_else(|| DecodeError::missing_required("TSP.Reference.identifier"))?,
        deprecated_type,
        deprecated_is_external,
    })
}

fn force_footnote_projection<'source>(
    view: &projection::FootnoteReferenceAttachmentArchiveLazyView<'source>,
) -> Result<FootnoteReferenceSnapshot<'source>, DecodeError> {
    let textual = view
        .super_
        .get()
        .map_err(DecodeError::from)?
        .map(|value| TextualSnapshot {
            string_equivalent: value.string_equivalent,
            kind: value.kind,
        });
    let contained_storage = view
        .contained_storage
        .get()
        .map_err(DecodeError::from)?
        .map(|value| force_reference_projection(&value))
        .transpose()?;
    Ok(FootnoteReferenceSnapshot {
        super_string_equivalent: textual.and_then(|value| value.string_equivalent),
        super_kind: textual.and_then(|value| value.kind),
        contained_storage,
        custom_mark_string: view.custom_mark_string,
    })
}

fn force_reference_projection(
    view: &projection::ReferenceLazyView<'_>,
) -> Result<ReferenceSnapshot, DecodeError> {
    if !view.has_identifier() {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    Ok(ReferenceSnapshot {
        identifier: NonZeroU64::new(view.identifier)
            .ok_or_else(|| DecodeError::zero_identifier("TSP.Reference.identifier"))?,
        deprecated_type: view.deprecated_type,
        deprecated_is_external: view.deprecated_is_external,
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
fn decode_int32(value: u64) -> i32 {
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
            take_exact(source, 8)?;
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
            let child_limit = recursion_limit.checked_sub(1).ok_or_else(|| {
                DecodeError::recursion_limit(recursion_limit.saturating_add(1), recursion_limit)
            })?;
            skip_strict_group(source, field_number, child_limit, budget)?;
            StrictValue::Group
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            take_exact(source, 4)?;
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
    reason = "Focused codec tests use explicit fixture panic messages."
)]
mod tests {
    use prost::Message as _;

    use super::*;
    use crate::{tsp, tswp};

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            64,
            source.len().saturating_mul(16).max(1),
            8,
        )
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            deprecated_type: Some(-7),
            deprecated_is_external: Some(false),
        }
    }

    fn fixture() -> tswp::FootnoteReferenceAttachmentArchive {
        tswp::FootnoteReferenceAttachmentArchive {
            super_: Some(tswp::TextualAttachmentArchive {
                string_equivalent: Some("*️".to_owned()),
                kind: Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32),
            }),
            contained_storage: Some(reference(42)),
            custom_mark_string: Some("*".to_owned()),
        }
    }

    #[test]
    fn canonical_fixture_matches_private_projection() -> Result<(), Box<dyn std::error::Error>> {
        let source = fixture().encode_to_vec();
        let snapshot = decode_footnote_reference(&source, options(&source))?;
        assert_eq!(snapshot.super_kind(), Some(2));
        assert_eq!(snapshot.super_string_equivalent(), Some("*️"));
        assert_eq!(snapshot.contained_storage().unwrap().identifier().get(), 42);
        assert_eq!(snapshot.custom_mark_string(), Some("*"));
        Ok(())
    }

    #[test]
    fn optional_edges_preserve_absence() -> Result<(), Box<dyn std::error::Error>> {
        let source = [0x0a, 0x00];
        let snapshot = decode_footnote_reference(&source, options(&source))?;
        assert_eq!(snapshot.super_kind(), None);
        assert_eq!(snapshot.contained_storage(), None);
        assert_eq!(snapshot.custom_mark_string(), None);
        Ok(())
    }

    #[test]
    fn root_and_one_nested_message_fit_a_one_level_profile()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = fixture().encode_to_vec();
        let snapshot = decode_footnote_reference(
            &source,
            DecodeOptions::new(source.len(), 64, source.len() * 16, 1),
        )?;
        assert_eq!(snapshot.contained_storage().unwrap().identifier().get(), 42);
        Ok(())
    }

    #[test]
    fn unknown_fields_remain_opaque_but_strictly_framed() -> Result<(), Box<dyn std::error::Error>>
    {
        let nested = [0x08, 0x2a, 0xa0, 0x06, 0x01];
        let mut source = vec![0x12, nested.len() as u8];
        source.extend_from_slice(&nested);
        source.extend_from_slice(&[0x98, 0x06, 0x01]);
        let snapshot = decode_footnote_reference(&source, options(&source))?;
        assert_eq!(snapshot.contained_storage().unwrap().identifier().get(), 42);
        Ok(())
    }

    #[test]
    fn malformed_known_fields_are_rejected_before_lazy_projection() {
        let mut duplicate = fixture().encode_to_vec();
        duplicate.extend_from_slice(&[0x18, 0x01, 0x18, 0x02]);
        assert_eq!(
            decode_footnote_reference(&duplicate, options(&duplicate))
                .expect_err("duplicate custom marker")
                .duplicate_singular_field(),
            Some("TSWP.FootnoteReferenceAttachmentArchive.custom_mark_string")
        );

        let mut bad_utf8 = fixture().encode_to_vec();
        bad_utf8.extend_from_slice(&[0x18, 0x01, 0xff]);
        assert!(decode_footnote_reference(&bad_utf8, options(&bad_utf8)).is_err());

        let zero_reference = [0x12, 0x02, 0x08, 0x00];
        let error = decode_footnote_reference(&zero_reference, options(&zero_reference))
            .expect_err("zero contained-storage identity");
        assert_eq!(
            error.zero_identifier_field(),
            Some("TSP.Reference.identifier")
        );
    }

    #[test]
    fn exact_limits_are_enforced() {
        let source = fixture().encode_to_vec();
        assert_eq!(
            decode_footnote_reference(
                &source,
                DecodeOptions::new(source.len() - 1, 64, source.len() * 16, 8)
            )
            .expect_err("byte ceiling")
            .resource_limit(),
            Some(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1,
            })
        );
        assert!(matches!(
            decode_footnote_reference(
                &source,
                DecodeOptions::new(source.len(), 1, source.len() * 16, 8)
            )
            .expect_err("field ceiling")
            .resource_limit(),
            Some(DecodeLimit::Fields {
                observed: 2,
                maximum: 1,
            })
        ));
        assert!(matches!(
            decode_footnote_reference(
                &source,
                DecodeOptions::new(source.len(), 64, source.len(), 8)
            )
            .expect_err("work ceiling")
            .resource_limit(),
            Some(DecodeLimit::Work {
                observed: _,
                maximum,
            }) if maximum == source.len()
        ));
        assert_eq!(
            decode_footnote_reference(
                &source,
                DecodeOptions::new(source.len(), 64, source.len() * 16, 0)
            )
            .expect_err("nesting ceiling")
            .resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 0,
                maximum: MAX_RECURSION_LIMIT,
            })
        );
    }
}
