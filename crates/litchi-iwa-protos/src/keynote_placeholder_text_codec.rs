//! Strict private Buffa projection for Keynote placeholder text owners.
//!
//! A bounded raw-wire pass rejects ambiguous or non-canonical selected fields
//! before Buffa observes them. Buffa then supplies a borrowed lazy-view
//! cross-check for the required placeholder/shape inheritance envelopes and
//! optional owned-storage edge. Unknown fields remain opaque in caller-owned
//! source bytes and are never retained or re-encoded by this module.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict semantic preflight intentionally precedes its low-level wire reader."
)]

use std::{fmt, num::NonZeroU64};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_placeholder_text_generated::LitchiIwaProjection as projection;

const PLACEHOLDER_SUPER_FIELD: u32 = 1;
const PLACEHOLDER_KIND_FIELD: u32 = 2;
const SHAPE_INFO_SUPER_FIELD: u32 = 1;
const SHAPE_INFO_OWNED_STORAGE_FIELD: u32 = 4;
const SHAPE_SUPER_FIELD: u32 = 1;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Explicit finite resource policy for one focused placeholder projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    initial_recursion_limit: u32,
}

impl DecodeOptions {
    /// Build a finite bytes/fields/work/nesting policy.
    ///
    /// Work charges strict preflight plus Buffa access for the root and every
    /// selected nested message that Buffa is forced to decode.
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
            initial_recursion_limit: recursion_limit,
        }
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }

    fn descend(self) -> Result<Self, DecodeError> {
        let recursion_limit = self.recursion_limit.checked_sub(1).ok_or_else(|| {
            DecodeError::recursion_limit_exceeded(
                self.initial_recursion_limit.saturating_add(1),
                self.initial_recursion_limit,
            )
        })?;
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

    /// Deprecated native type hint, retained only for projection parity.
    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }

    /// Deprecated external-reference marker, retained for projection parity.
    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }
}

/// Recognized values of the optional native placeholder-kind hint.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub enum PlaceholderKind {
    /// Generic placeholder or the field's declared default.
    Generic = 0,
    /// Slide-number placeholder.
    SlideNumber = 1,
    /// Title placeholder.
    Title = 2,
    /// Body placeholder.
    Body = 3,
    /// Object placeholder.
    Object = 4,
}

impl PlaceholderKind {
    const fn from_raw(value: i32) -> Option<Self> {
        match value {
            0 => Some(Self::Generic),
            1 => Some(Self::SlideNumber),
            2 => Some(Self::Title),
            3 => Some(Self::Body),
            4 => Some(Self::Object),
            _ => None,
        }
    }
}

/// Generated-free owner facts from one `KN.PlaceholderArchive`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PlaceholderTextOwnerSnapshot {
    owned_storage: Option<ReferenceSnapshot>,
    kind: Option<i32>,
}

impl PlaceholderTextOwnerSnapshot {
    /// Optional shape-owned `TSWP.StorageArchive` edge.
    #[must_use]
    pub const fn owned_storage(self) -> Option<ReferenceSnapshot> {
        self.owned_storage
    }

    /// Exact explicitly encoded placeholder kind, excluding the schema default.
    #[must_use]
    pub const fn kind(self) -> Option<i32> {
        self.kind
    }

    /// Recognized explicitly encoded placeholder kind.
    ///
    /// `None` means either that the optional field was absent or its numeric
    /// value is unknown to the checked-in schema. Call [`Self::kind`] when
    /// those cases must be distinguished.
    #[must_use]
    pub const fn recognized_kind(self) -> Option<PlaceholderKind> {
        match self.kind {
            Some(value) => PlaceholderKind::from_raw(value),
            None => None,
        }
    }
}

/// Failure from strict placeholder-owner preflight or Buffa cross-check.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

/// Resource limit reported by [`DecodeError::limit_kind`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecodeLimitKind {
    /// A complete protobuf payload exceeded its configured byte ceiling.
    MessageBytes,
    /// Protobuf nesting exceeded its configured depth ceiling.
    Recursion,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    ZeroIdentifier(&'static str),
    MessageByteLimit { observed: usize, maximum: usize },
    RecursionLimit { observed: u32, maximum: u32 },
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
    Projection,
}

impl DecodeError {
    const fn message_byte_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::MessageByteLimit { observed, maximum },
        }
    }

    const fn recursion_limit_exceeded(observed: u32, maximum: u32) -> Self {
        Self {
            kind: DecodeErrorKind::RecursionLimit { observed, maximum },
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

    /// Required schema field absent from source, when applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        let DecodeErrorKind::MissingRequired(field) = self.kind else {
            return None;
        };
        Some(field)
    }

    /// Repeated singular field, when applicable.
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

    /// Resource-limit category, without exposing source content.
    #[must_use]
    pub const fn limit_kind(&self) -> Option<DecodeLimitKind> {
        match &self.kind {
            DecodeErrorKind::MessageByteLimit { .. }
            | DecodeErrorKind::Wire(buffa::DecodeError::MessageTooLarge) => {
                Some(DecodeLimitKind::MessageBytes)
            },
            DecodeErrorKind::RecursionLimit { .. }
            | DecodeErrorKind::Wire(buffa::DecodeError::RecursionLimitExceeded) => {
                Some(DecodeLimitKind::Recursion)
            },
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Observed and configured message-byte ceiling, when this codec measured
    /// the failure before entering Buffa.
    #[must_use]
    pub const fn message_byte_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::MessageByteLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Observed and configured recursion ceiling, when strict preflight
    /// measured the failure before entering Buffa.
    #[must_use]
    pub const fn recursion_limit_values(&self) -> Option<(u32, u32)> {
        let DecodeErrorKind::RecursionLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
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
            DecodeErrorKind::ZeroIdentifier(field) => write!(formatter, "{field} is zero"),
            DecodeErrorKind::MessageByteLimit { observed, maximum } => write!(
                formatter,
                "Keynote placeholder projection received {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::RecursionLimit { observed, maximum } => write!(
                formatter,
                "Keynote placeholder projection requires nesting depth {observed}; maximum is {maximum}"
            ),
            DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Keynote placeholder projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Keynote placeholder projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Projection => formatter.write_str(
                "Keynote placeholder strict preflight disagrees with the Buffa projection",
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

/// Decode text ownership from one complete `KN.PlaceholderArchive`.
///
/// Strict preflight proves the complete required inheritance chain and every
/// selected nested reference before Buffa's deferred values are forced. No
/// generated value or preservation state escapes this function.
pub fn decode_placeholder_text_owner(
    source: &[u8],
    options: DecodeOptions,
) -> Result<PlaceholderTextOwnerSnapshot, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_placeholder(source, options, &mut budget)?;
    let view: projection::PlaceholderArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    let projected = force_placeholder_projection(&view)?;
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

/// Decode only the optional placeholder-to-storage object identifier.
///
/// The convenience retains complete inheritance-envelope validation.
pub fn decode_placeholder_storage_reference(
    source: &[u8],
    options: DecodeOptions,
) -> Result<Option<u64>, DecodeError> {
    Ok(decode_placeholder_text_owner(source, options)?
        .owned_storage()
        .map(|reference| reference.identifier().get()))
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let max_buffa_message_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
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
        return Err(DecodeError::recursion_limit_exceeded(1, 0));
    }
    if options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::recursion_limit_exceeded(
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
        let strict_and_projection = bytes.saturating_mul(2);
        let observed = self.work_bytes.saturating_add(strict_and_projection);
        if observed > self.max_work_bytes {
            return Err(DecodeError::work_limit(observed, self.max_work_bytes));
        }
        self.work_bytes = observed;
        Ok(())
    }

    const fn recursion_limit_exceeded(&self) -> DecodeError {
        DecodeError::recursion_limit_exceeded(
            self.max_recursion_limit.saturating_add(1),
            self.max_recursion_limit,
        )
    }
}

fn preflight_placeholder(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<PlaceholderTextOwnerSnapshot, DecodeError> {
    budget.charge_message(source.len())?;
    let nested = options.descend()?;
    let mut shape_info = None;
    let mut kind = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            PLACEHOLDER_SUPER_FIELD => {
                if shape_info.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "KN.PlaceholderArchive.super",
                    ));
                }
                shape_info = Some(preflight_shape_info(
                    field.length_delimited()?,
                    nested,
                    budget,
                )?);
            },
            PLACEHOLDER_KIND_FIELD => {
                if kind.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "KN.PlaceholderArchive.kind",
                    ));
                }
                kind = Some(decode_int32(require_canonical_int32(field.varint()?)?));
            },
            _ => {},
        }
    }
    Ok(PlaceholderTextOwnerSnapshot {
        owned_storage: shape_info
            .ok_or_else(|| DecodeError::missing_required("KN.PlaceholderArchive.super"))?,
        kind,
    })
}

fn preflight_shape_info(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Option<ReferenceSnapshot>, DecodeError> {
    budget.charge_message(source.len())?;
    let nested = options.descend()?;
    let mut saw_super = false;
    let mut owned_storage = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            SHAPE_INFO_SUPER_FIELD => {
                if saw_super {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.ShapeInfoArchive.super",
                    ));
                }
                saw_super = true;
                preflight_shape(field.length_delimited()?, nested, budget)?;
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
            _ => {},
        }
    }
    if !saw_super {
        return Err(DecodeError::missing_required("TSWP.ShapeInfoArchive.super"));
    }
    Ok(owned_storage)
}

fn preflight_shape(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_message(source.len())?;
    let nested = options.descend()?;
    let mut saw_super = false;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        if field.number != SHAPE_SUPER_FIELD {
            continue;
        }
        if saw_super {
            return Err(DecodeError::duplicate_singular("TSD.ShapeArchive.super"));
        }
        saw_super = true;
        preflight_drawable(field.length_delimited()?, nested, budget)?;
    }
    if !saw_super {
        return Err(DecodeError::missing_required("TSD.ShapeArchive.super"));
    }
    Ok(())
}

fn preflight_drawable(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_message(source.len())?;
    let mut remaining = source;
    while next_strict_field(&mut remaining, options, budget)?.is_some() {}
    Ok(())
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

fn force_placeholder_projection(
    view: &projection::PlaceholderArchiveLazyView<'_>,
) -> Result<PlaceholderTextOwnerSnapshot, DecodeError> {
    let shape_info = view
        .super_
        .get()?
        .ok_or_else(|| DecodeError::missing_required("KN.PlaceholderArchive.super"))?;
    let shape = shape_info
        .super_
        .get()?
        .ok_or_else(|| DecodeError::missing_required("TSWP.ShapeInfoArchive.super"))?;
    let _drawable = shape
        .super_
        .get()?
        .ok_or_else(|| DecodeError::missing_required("TSD.ShapeArchive.super"))?;
    let owned_storage = shape_info
        .owned_storage
        .get()?
        .map(|reference| force_reference_projection(&reference))
        .transpose()?;
    Ok(PlaceholderTextOwnerSnapshot {
        owned_storage,
        kind: view.kind,
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
                .ok_or_else(|| budget.recursion_limit_exceeded())?;
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
    reason = "Focused negative tests use explicit panic messages."
)]
mod tests {
    use prost::Message as _;

    use super::{
        DecodeLimitKind, DecodeOptions, PlaceholderKind, decode_placeholder_storage_reference,
        decode_placeholder_text_owner,
    };
    use crate::{kn, tsd, tsp, tswp};

    fn options(source: &[u8], recursion_limit: u32) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            source.len().max(1),
            source.len().saturating_mul(12).max(1),
            recursion_limit,
        )
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            deprecated_type: Some(-7),
            deprecated_is_external: Some(false),
        }
    }

    fn placeholder(storage: Option<u64>, kind: Option<i32>) -> kn::PlaceholderArchive {
        kn::PlaceholderArchive {
            super_: tswp::ShapeInfoArchive {
                super_: tsd::ShapeArchive {
                    super_: tsd::DrawableArchive::default(),
                    ..Default::default()
                },
                owned_storage: storage.map(reference),
                ..Default::default()
            },
            kind,
        }
    }

    #[test]
    fn canonical_prost_owner_matches_lazy_projection() -> Result<(), Box<dyn std::error::Error>> {
        let source = placeholder(Some(42), Some(2)).encode_to_vec();
        let owner = decode_placeholder_text_owner(&source, options(&source, 4))?;
        assert_eq!(
            owner
                .owned_storage()
                .map(|reference| reference.identifier().get()),
            Some(42)
        );
        assert_eq!(owner.kind(), Some(2));
        assert_eq!(owner.recognized_kind(), Some(PlaceholderKind::Title));
        assert_eq!(
            decode_placeholder_storage_reference(&source, options(&source, 4))?,
            Some(42)
        );
        Ok(())
    }

    #[test]
    fn absent_optional_storage_and_kind_are_preserved() -> Result<(), Box<dyn std::error::Error>> {
        let source = placeholder(None, None).encode_to_vec();
        let owner = decode_placeholder_text_owner(&source, options(&source, 4))?;
        assert_eq!(owner.owned_storage(), None);
        assert_eq!(owner.kind(), None);
        assert_eq!(owner.recognized_kind(), None);
        Ok(())
    }

    #[test]
    fn unknown_kind_remains_distinguishable_from_absence() -> Result<(), Box<dyn std::error::Error>>
    {
        let source = placeholder(Some(42), Some(99)).encode_to_vec();
        let owner = decode_placeholder_text_owner(&source, options(&source, 4))?;
        assert_eq!(owner.kind(), Some(99));
        assert_eq!(owner.recognized_kind(), None);
        Ok(())
    }

    #[test]
    fn required_inheritance_envelopes_and_nonzero_ids_are_enforced() {
        assert_eq!(
            decode_placeholder_text_owner(&[], options(&[], 4))
                .expect_err("missing placeholder super")
                .missing_required_field(),
            Some("KN.PlaceholderArchive.super")
        );
        assert_eq!(
            decode_placeholder_text_owner(&[0x0a, 0x00], options(&[0x0a, 0x00], 4))
                .expect_err("missing shape-info super")
                .missing_required_field(),
            Some("TSWP.ShapeInfoArchive.super")
        );
        let missing_drawable = [0x0a, 0x02, 0x0a, 0x00];
        assert_eq!(
            decode_placeholder_text_owner(&missing_drawable, options(&missing_drawable, 4))
                .expect_err("missing shape super")
                .missing_required_field(),
            Some("TSD.ShapeArchive.super")
        );
        let zero_storage = [0x0a, 0x08, 0x0a, 0x02, 0x0a, 0x00, 0x22, 0x02, 0x08, 0x00];
        assert_eq!(
            decode_placeholder_text_owner(&zero_storage, options(&zero_storage, 4))
                .expect_err("zero storage reference")
                .zero_identifier_field(),
            Some("TSP.Reference.identifier")
        );
    }

    #[test]
    fn duplicate_selected_and_nested_fields_fail_before_buffa_last_wins() {
        let mut duplicate_kind = placeholder(None, Some(2)).encode_to_vec();
        duplicate_kind.extend_from_slice(&[0x10, 0x03]);
        assert_eq!(
            decode_placeholder_text_owner(&duplicate_kind, options(&duplicate_kind, 4))
                .expect_err("duplicate kind")
                .duplicate_singular_field(),
            Some("KN.PlaceholderArchive.kind")
        );

        let duplicate_identifier = [
            0x0a, 0x0a, 0x0a, 0x02, 0x0a, 0x00, 0x22, 0x04, 0x08, 0x01, 0x08, 0x02,
        ];
        assert_eq!(
            decode_placeholder_text_owner(
                &duplicate_identifier,
                options(&duplicate_identifier, 4),
            )
            .expect_err("duplicate reference identifier")
            .duplicate_singular_field(),
            Some("TSP.Reference.identifier")
        );
    }

    #[test]
    fn noncanonical_fields_are_rejected() {
        let mut overlong_unknown = placeholder(None, None).encode_to_vec();
        overlong_unknown.extend_from_slice(&[0xa0, 0x00, 0x01]);
        assert_eq!(
            decode_placeholder_text_owner(&overlong_unknown, options(&overlong_unknown, 4))
                .expect_err("overlong unknown key")
                .noncanonical_reason(),
            Some("protobuf field key")
        );

        let mut bad_kind = placeholder(None, None).encode_to_vec();
        bad_kind.extend_from_slice(&[0x10, 0x80, 0x80, 0x80, 0x80, 0x08]);
        assert_eq!(
            decode_placeholder_text_owner(&bad_kind, options(&bad_kind, 4))
                .expect_err("noncanonical int32")
                .noncanonical_reason(),
            Some("int32 scalar is not a sign-extended 32-bit value")
        );
    }

    #[test]
    fn unknown_fields_remain_opaque_but_strictly_framed() -> Result<(), Box<dyn std::error::Error>>
    {
        let mut source = placeholder(Some(42), Some(3)).encode_to_vec();
        source.extend_from_slice(&[0x82, 0x02, 0x02, 0xff, 0x00]);
        let owner = decode_placeholder_text_owner(&source, options(&source, 4))?;
        assert_eq!(owner.recognized_kind(), Some(PlaceholderKind::Body));
        Ok(())
    }

    #[test]
    fn exact_resource_limits_pass_and_one_less_fails() -> Result<(), Box<dyn std::error::Error>> {
        let source = placeholder(Some(42), Some(2)).encode_to_vec();
        let expected = decode_placeholder_text_owner(&source, options(&source, 4))?;
        let byte_error = decode_placeholder_text_owner(
            &source,
            DecodeOptions::new(source.len() - 1, source.len(), source.len() * 12, 4),
        )
        .expect_err("message-byte limit");
        assert_eq!(byte_error.limit_kind(), Some(DecodeLimitKind::MessageBytes));
        assert_eq!(
            byte_error.message_byte_limit_values(),
            Some((source.len(), source.len() - 1))
        );

        let field_error = decode_placeholder_text_owner(
            &source,
            DecodeOptions::new(source.len(), 1, source.len() * 12, 4),
        )
        .expect_err("field limit");
        assert_eq!(field_error.field_limit_values(), Some((2, 1)));

        let exact_work = strict_work_for_placeholder(&source)?;
        assert_eq!(
            decode_placeholder_text_owner(
                &source,
                DecodeOptions::new(source.len(), source.len(), exact_work, 4),
            )?,
            expected
        );
        assert_eq!(
            decode_placeholder_text_owner(
                &source,
                DecodeOptions::new(source.len(), source.len(), exact_work - 1, 4),
            )
            .expect_err("work limit")
            .work_limit_values(),
            Some((exact_work, exact_work - 1))
        );
        let nesting_error = decode_placeholder_text_owner(
            &source,
            DecodeOptions::new(source.len(), source.len(), source.len() * 12, 2),
        )
        .expect_err("recursion limit");
        assert_eq!(nesting_error.limit_kind(), Some(DecodeLimitKind::Recursion));
        assert_eq!(nesting_error.recursion_limit_values(), Some((3, 2)));
        assert_eq!(nesting_error.message_byte_limit_values(), None);
        assert_eq!(nesting_error.field_limit_values(), None);
        Ok(())
    }

    #[test]
    fn invalid_resource_profiles_have_structured_limit_errors() {
        let source = placeholder(None, None).encode_to_vec();
        let zero_depth = decode_placeholder_text_owner(
            &source,
            DecodeOptions::new(source.len(), source.len(), source.len() * 12, 0),
        )
        .expect_err("zero recursion limit");
        assert_eq!(zero_depth.limit_kind(), Some(DecodeLimitKind::Recursion));
        assert_eq!(zero_depth.recursion_limit_values(), Some((1, 0)));

        let excessive_depth = decode_placeholder_text_owner(
            &source,
            DecodeOptions::new(source.len(), source.len(), source.len() * 12, 65),
        )
        .expect_err("recursion limit above codec ceiling");
        assert_eq!(excessive_depth.recursion_limit_values(), Some((65, 64)));
    }

    fn strict_work_for_placeholder(source: &[u8]) -> Result<usize, Box<dyn std::error::Error>> {
        let native = kn::PlaceholderArchive::decode(source)?;
        let shape_info = native.super_.encode_to_vec();
        let shape = native.super_.super_.encode_to_vec();
        let drawable = native.super_.super_.super_.encode_to_vec();
        let storage = native
            .super_
            .owned_storage
            .map(|reference| reference.encode_to_vec());
        Ok(source
            .len()
            .checked_add(shape_info.len())
            .and_then(|value| value.checked_add(shape.len()))
            .and_then(|value| value.checked_add(drawable.len()))
            .and_then(|value| value.checked_add(storage.as_ref().map_or(0, Vec::len)))
            .and_then(|value| value.checked_mul(2))
            .ok_or("test work overflow")?)
    }

    #[test]
    fn malformed_wire_never_panics() {
        let malformed: [&[u8]; 7] = [
            &[0x80],
            &[0x00],
            &[0x0f],
            &[0x0a, 0x02, 0x08],
            &[0x0b],
            &[0x0a, 0x01, 0x0a],
            &[0x10, 0x80],
        ];
        for source in malformed {
            assert!(decode_placeholder_text_owner(source, options(source, 4)).is_err());
        }
    }
}
