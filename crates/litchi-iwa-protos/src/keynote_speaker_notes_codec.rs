//! Strict private Buffa projection for Keynote speaker-note owners.
//!
//! A bounded raw-wire pass rejects ambiguous or non-canonical selected fields
//! before Buffa observes them. Buffa then supplies borrowed lazy-view
//! cross-checks for the slide's semantic note/title/body edges and the note's
//! storage edge. Unknown fields remain opaque in caller-owned source bytes and
//! are never retained or re-encoded by this module.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict semantic preflight intentionally precedes its low-level wire reader."
)]

use std::{fmt, num::NonZeroU64};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_speaker_notes_generated::LitchiIwaProjection as projection;

const SLIDE_STYLE_FIELD: u32 = 1;
const SLIDE_TRANSITION_FIELD: u32 = 4;
const SLIDE_TITLE_PLACEHOLDER_FIELD: u32 = 5;
const SLIDE_BODY_PLACEHOLDER_FIELD: u32 = 6;
const SLIDE_NAME_FIELD: u32 = 10;
const SLIDE_IN_DOCUMENT_FIELD: u32 = 19;
const SLIDE_NUMBER_PLACEHOLDER_FIELD: u32 = 20;
const SLIDE_NOTE_FIELD: u32 = 27;
const TRANSITION_ATTRIBUTES_FIELD: u32 = 2;
const NOTE_STORAGE_FIELD: u32 = 1;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Explicit finite resource policy for one focused speaker-note projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build a finite bytes/fields/work/nesting policy.
    ///
    /// Work charges both strict preflight and Buffa access for the root and
    /// every selected nested message that Buffa is forced to decode.
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

/// Borrowed generated-free owner facts from one `KN.SlideArchive`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideNotesOwnerSnapshot<'source> {
    style: ReferenceSnapshot,
    title_placeholder: Option<ReferenceSnapshot>,
    body_placeholder: Option<ReferenceSnapshot>,
    name: Option<&'source str>,
    in_document: bool,
    slide_number_placeholder: Option<ReferenceSnapshot>,
    note: Option<ReferenceSnapshot>,
}

impl<'source> SlideNotesOwnerSnapshot<'source> {
    /// Required slide-style edge, validated to preserve the slide envelope.
    #[must_use]
    pub const fn style(self) -> ReferenceSnapshot {
        self.style
    }

    /// Optional edge to the slide's semantic title placeholder.
    #[must_use]
    pub const fn title_placeholder(self) -> Option<ReferenceSnapshot> {
        self.title_placeholder
    }

    /// Optional edge to the slide's semantic body placeholder.
    #[must_use]
    pub const fn body_placeholder(self) -> Option<ReferenceSnapshot> {
        self.body_placeholder
    }

    /// Optional exact navigator name borrowed from source bytes.
    #[must_use]
    pub const fn name(self) -> Option<&'source str> {
        self.name
    }

    /// Required native in-document state.
    #[must_use]
    pub const fn in_document(self) -> bool {
        self.in_document
    }

    /// Optional edge to the slide-number placeholder owner.
    #[must_use]
    pub const fn slide_number_placeholder(self) -> Option<ReferenceSnapshot> {
        self.slide_number_placeholder
    }

    /// Optional edge to `KN.NoteArchive`.
    #[must_use]
    pub const fn note(self) -> Option<ReferenceSnapshot> {
        self.note
    }
}

/// Generated-free required storage edge from one `KN.NoteArchive`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NoteOwnerSnapshot {
    contained_storage: ReferenceSnapshot,
}

impl NoteOwnerSnapshot {
    /// Required edge to the speaker-note `TSWP.StorageArchive`.
    #[must_use]
    pub const fn contained_storage(self) -> ReferenceSnapshot {
        self.contained_storage
    }
}

/// Failure from strict speaker-note owner preflight or Buffa cross-check.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

/// A content-free wire-resource classification for [`DecodeError`].
///
/// Values describe only configured limits and aggregate byte counts; they
/// never expose decoded protobuf field values or caller-owned source bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WireResourceLimit {
    /// The input or configured message-byte ceiling could not be honored.
    ///
    /// `observed` is populated only when the complete input length was known
    /// at the point it exceeded `maximum`.
    Bytes {
        observed: Option<usize>,
        maximum: Option<usize>,
    },
    /// The configured or enforced protobuf nesting ceiling was exceeded.
    ///
    /// `observed` is populated only for an invalid configured ceiling; a
    /// decoder recursion failure does not reveal a trustworthy exact depth.
    Nesting {
        observed: Option<u32>,
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
    InvalidUtf8(&'static str),
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
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

    /// UTF-8 string field rejected during raw preflight, when applicable.
    #[must_use]
    pub const fn invalid_utf8_field(&self) -> Option<&'static str> {
        let DecodeErrorKind::InvalidUtf8(field) = self.kind else {
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

    /// Wire byte/nesting resource failure, independent of Buffa error text.
    ///
    /// This is intentionally content-free so format adapters can map trusted
    /// resource-policy failures without exposing raw decoder internals.
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
                formatter.write_str("Keynote speaker-note wire byte limit exceeded")
            },
            DecodeErrorKind::WireResourceLimit(WireResourceLimit::Nesting { .. }) => {
                formatter.write_str("Keynote speaker-note wire nesting limit exceeded")
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
            DecodeErrorKind::InvalidUtf8(field) => write!(formatter, "{field} is invalid UTF-8"),
            DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Keynote speaker-note projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Keynote speaker-note projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Projection => formatter.write_str(
                "Keynote speaker-note strict preflight disagrees with the Buffa projection",
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

/// Decode slide speaker-note ownership from one complete `KN.SlideArchive`.
///
/// Strict preflight proves the required slide envelope and every selected
/// nested reference before Buffa's deferred values are forced. The returned
/// name borrows `source`; no generated value or preservation state escapes.
pub fn decode_slide_notes_owner<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<SlideNotesOwnerSnapshot<'source>, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_slide(source, options, &mut budget)
        .map_err(|error| error.with_recursion_limit_context(options.recursion_limit))?;
    let view: projection::SlideArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let projected = force_slide_projection(&view)
        .map_err(|error| error.with_recursion_limit_context(options.recursion_limit))?;
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

/// Decode the required storage edge from one complete `KN.NoteArchive`.
pub fn decode_note_owner(
    source: &[u8],
    options: DecodeOptions,
) -> Result<NoteOwnerSnapshot, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_note(source, options, &mut budget)
        .map_err(|error| error.with_recursion_limit_context(options.recursion_limit))?;
    let view: projection::NoteArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let projected = force_note_projection(&view)
        .map_err(|error| error.with_recursion_limit_context(options.recursion_limit))?;
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

/// Decode only the optional slide-to-note object identifier.
///
/// This generated-free convenience retains the complete strict slide-envelope
/// validation performed by [`decode_slide_notes_owner`].
pub fn decode_slide_note_reference(
    source: &[u8],
    options: DecodeOptions,
) -> Result<Option<u64>, DecodeError> {
    Ok(decode_slide_notes_owner(source, options)?
        .note()
        .map(|reference| reference.identifier().get()))
}

/// Decode only the required note-to-storage object identifier.
///
/// The identifier is proven non-zero by strict preflight before conversion to
/// the compatibility scalar returned to the format-owned package adapter.
pub fn decode_note_storage_reference(
    source: &[u8],
    options: DecodeOptions,
) -> Result<u64, DecodeError> {
    Ok(decode_note_owner(source, options)?
        .contained_storage()
        .identifier()
        .get())
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

fn preflight_slide<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<SlideNotesOwnerSnapshot<'source>, DecodeError> {
    budget.charge_message(source.len())?;
    let nested = options.descend()?;
    let mut style = None;
    let mut saw_transition = false;
    let mut title_placeholder = None;
    let mut body_placeholder = None;
    let mut name = None;
    let mut in_document = None;
    let mut slide_number_placeholder = None;
    let mut note = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            SLIDE_STYLE_FIELD => {
                if style.is_some() {
                    return Err(DecodeError::duplicate_singular("KN.SlideArchive.style"));
                }
                style = Some(preflight_reference(
                    field.length_delimited()?,
                    nested,
                    budget,
                )?);
            },
            SLIDE_TRANSITION_FIELD => {
                if saw_transition {
                    return Err(DecodeError::duplicate_singular(
                        "KN.SlideArchive.transition",
                    ));
                }
                saw_transition = true;
                preflight_transition(field.length_delimited()?, nested, budget)?;
            },
            SLIDE_TITLE_PLACEHOLDER_FIELD => {
                if title_placeholder.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "KN.SlideArchive.titlePlaceholder",
                    ));
                }
                title_placeholder = Some(preflight_reference(
                    field.length_delimited()?,
                    nested,
                    budget,
                )?);
            },
            SLIDE_BODY_PLACEHOLDER_FIELD => {
                if body_placeholder.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "KN.SlideArchive.bodyPlaceholder",
                    ));
                }
                body_placeholder = Some(preflight_reference(
                    field.length_delimited()?,
                    nested,
                    budget,
                )?);
            },
            SLIDE_NAME_FIELD => {
                if name.is_some() {
                    return Err(DecodeError::duplicate_singular("KN.SlideArchive.name"));
                }
                name = Some(
                    std::str::from_utf8(field.length_delimited()?)
                        .map_err(|_error| DecodeError::invalid_utf8("KN.SlideArchive.name"))?,
                );
            },
            SLIDE_IN_DOCUMENT_FIELD => {
                if in_document.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "KN.SlideArchive.inDocument",
                    ));
                }
                in_document = Some(require_canonical_bool(field.varint()?)?);
            },
            SLIDE_NUMBER_PLACEHOLDER_FIELD => {
                if slide_number_placeholder.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "KN.SlideArchive.slideNumberPlaceholder",
                    ));
                }
                slide_number_placeholder = Some(preflight_reference(
                    field.length_delimited()?,
                    nested,
                    budget,
                )?);
            },
            SLIDE_NOTE_FIELD => {
                if note.is_some() {
                    return Err(DecodeError::duplicate_singular("KN.SlideArchive.note"));
                }
                note = Some(preflight_reference(
                    field.length_delimited()?,
                    nested,
                    budget,
                )?);
            },
            _ => {},
        }
    }
    if !saw_transition {
        return Err(DecodeError::missing_required("KN.SlideArchive.transition"));
    }
    Ok(SlideNotesOwnerSnapshot {
        style: style.ok_or_else(|| DecodeError::missing_required("KN.SlideArchive.style"))?,
        title_placeholder,
        body_placeholder,
        name,
        in_document: in_document
            .ok_or_else(|| DecodeError::missing_required("KN.SlideArchive.inDocument"))?,
        slide_number_placeholder,
        note,
    })
}

fn preflight_transition(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_message(source.len())?;
    let nested = options.descend()?;
    let mut saw_attributes = false;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        if field.number != TRANSITION_ATTRIBUTES_FIELD {
            continue;
        }
        if saw_attributes {
            return Err(DecodeError::duplicate_singular(
                "KN.TransitionArchive.attributes",
            ));
        }
        saw_attributes = true;
        preflight_opaque_message(field.length_delimited()?, nested, budget)?;
    }
    if !saw_attributes {
        return Err(DecodeError::missing_required(
            "KN.TransitionArchive.attributes",
        ));
    }
    Ok(())
}

fn preflight_opaque_message(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_message(source.len())?;
    let mut remaining = source;
    while next_strict_field(&mut remaining, options, budget)?.is_some() {}
    Ok(())
}

fn preflight_note(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<NoteOwnerSnapshot, DecodeError> {
    budget.charge_message(source.len())?;
    let nested = options.descend()?;
    let mut contained_storage = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        if field.number != NOTE_STORAGE_FIELD {
            continue;
        }
        if contained_storage.is_some() {
            return Err(DecodeError::duplicate_singular(
                "KN.NoteArchive.containedStorage",
            ));
        }
        contained_storage = Some(preflight_reference(
            field.length_delimited()?,
            nested,
            budget,
        )?);
    }
    Ok(NoteOwnerSnapshot {
        contained_storage: contained_storage
            .ok_or_else(|| DecodeError::missing_required("KN.NoteArchive.containedStorage"))?,
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

fn force_slide_projection<'source>(
    view: &projection::SlideArchiveLazyView<'source>,
) -> Result<SlideNotesOwnerSnapshot<'source>, DecodeError> {
    let style_view = view
        .style
        .get()?
        .ok_or_else(|| DecodeError::missing_required("KN.SlideArchive.style"))?;
    let style = force_reference_projection(&style_view)?;
    let transition = view
        .transition
        .get()?
        .ok_or_else(|| DecodeError::missing_required("KN.SlideArchive.transition"))?;
    let _attributes = transition
        .attributes
        .get()?
        .ok_or_else(|| DecodeError::missing_required("KN.TransitionArchive.attributes"))?;
    if !view.has_in_document() {
        return Err(DecodeError::missing_required("KN.SlideArchive.inDocument"));
    }
    let note = view
        .note
        .get()?
        .map(|note| force_reference_projection(&note))
        .transpose()?;
    let title_placeholder = view
        .title_placeholder
        .get()?
        .map(|reference| force_reference_projection(&reference))
        .transpose()?;
    let body_placeholder = view
        .body_placeholder
        .get()?
        .map(|reference| force_reference_projection(&reference))
        .transpose()?;
    let slide_number_placeholder = view
        .slide_number_placeholder
        .get()?
        .map(|reference| force_reference_projection(&reference))
        .transpose()?;
    Ok(SlideNotesOwnerSnapshot {
        style,
        title_placeholder,
        body_placeholder,
        name: view.name,
        in_document: view.in_document,
        slide_number_placeholder,
        note,
    })
}

fn force_note_projection(
    view: &projection::NoteArchiveLazyView<'_>,
) -> Result<NoteOwnerSnapshot, DecodeError> {
    let contained_storage = view
        .contained_storage
        .get()?
        .ok_or_else(|| DecodeError::missing_required("KN.NoteArchive.containedStorage"))?;
    Ok(NoteOwnerSnapshot {
        contained_storage: force_reference_projection(&contained_storage)?,
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
    use prost::Message as _;

    use super::{
        DecodeOptions, NoteOwnerSnapshot, SlideNotesOwnerSnapshot, WireResourceLimit,
        decode_note_owner, decode_note_storage_reference, decode_slide_note_reference,
        decode_slide_notes_owner,
    };
    use crate::{kn, tsp};

    fn options(source: &[u8], recursion_limit: u32) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            source.len().max(1),
            source.len().saturating_mul(8).max(1),
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

    fn slide(name: Option<&str>, note: Option<u64>) -> kn::SlideArchive {
        kn::SlideArchive {
            style: reference(11),
            transition: kn::TransitionArchive {
                attributes: kn::TransitionAttributesArchive::default(),
            },
            name: name.map(str::to_owned),
            in_document: true,
            note: note.map(reference),
            ..Default::default()
        }
    }

    fn decode_slide(source: &[u8]) -> Result<SlideNotesOwnerSnapshot<'_>, super::DecodeError> {
        decode_slide_notes_owner(source, options(source, 3))
    }

    fn decode_note(source: &[u8]) -> Result<NoteOwnerSnapshot, super::DecodeError> {
        decode_note_owner(source, options(source, 2))
    }

    #[test]
    fn canonical_prost_slide_and_note_match_borrowed_projection()
    -> Result<(), Box<dyn std::error::Error>> {
        let mut native_slide = slide(Some("Agenda 🚀"), Some(42));
        native_slide.title_placeholder = Some(reference(51));
        native_slide.body_placeholder = Some(reference(52));
        let slide_source = native_slide.encode_to_vec();
        let owner = decode_slide(&slide_source)?;
        assert_eq!(owner.style().identifier().get(), 11);
        assert_eq!(owner.style().deprecated_type(), Some(-7));
        assert_eq!(
            owner
                .title_placeholder()
                .map(|reference| reference.identifier().get()),
            Some(51)
        );
        assert_eq!(
            owner
                .body_placeholder()
                .map(|reference| reference.identifier().get()),
            Some(52)
        );
        assert_eq!(owner.name(), Some("Agenda 🚀"));
        assert!(owner.in_document());
        assert_eq!(
            owner.note().map(|reference| reference.identifier().get()),
            Some(42)
        );

        let note_source = kn::NoteArchive {
            contained_storage: reference(99),
        }
        .encode_to_vec();
        let note = decode_note(&note_source)?;
        assert_eq!(note.contained_storage().identifier().get(), 99);
        assert_eq!(note.contained_storage().deprecated_type(), Some(-7));
        assert_eq!(
            decode_slide_note_reference(&slide_source, options(&slide_source, 3))?,
            Some(42)
        );
        assert_eq!(
            decode_note_storage_reference(&note_source, options(&note_source, 2))?,
            99
        );
        Ok(())
    }

    #[test]
    fn absent_optional_name_and_note_are_preserved() -> Result<(), Box<dyn std::error::Error>> {
        let source = slide(None, None).encode_to_vec();
        let owner = decode_slide(&source)?;
        assert_eq!(owner.title_placeholder(), None);
        assert_eq!(owner.body_placeholder(), None);
        assert_eq!(owner.name(), None);
        assert_eq!(owner.note(), None);
        Ok(())
    }

    #[test]
    fn required_envelopes_and_nonzero_identifiers_are_enforced() {
        let missing_style = [0x22, 0x02, 0x12, 0x00, 0x98, 0x01, 0x01];
        assert_eq!(
            decode_slide(&missing_style)
                .expect_err("missing style")
                .missing_required_field(),
            Some("KN.SlideArchive.style")
        );

        let missing_transition = [0x0a, 0x02, 0x08, 0x01, 0x98, 0x01, 0x01];
        assert_eq!(
            decode_slide(&missing_transition)
                .expect_err("missing transition")
                .missing_required_field(),
            Some("KN.SlideArchive.transition")
        );

        let missing_attributes = [0x0a, 0x02, 0x08, 0x01, 0x22, 0x00, 0x98, 0x01, 0x01];
        assert_eq!(
            decode_slide(&missing_attributes)
                .expect_err("missing transition attributes")
                .missing_required_field(),
            Some("KN.TransitionArchive.attributes")
        );

        let zero_style = [
            0x0a, 0x02, 0x08, 0x00, 0x22, 0x02, 0x12, 0x00, 0x98, 0x01, 0x01,
        ];
        assert_eq!(
            decode_slide(&zero_style)
                .expect_err("zero style")
                .zero_identifier_field(),
            Some("TSP.Reference.identifier")
        );

        assert_eq!(
            decode_note(&[0x0a, 0x00])
                .expect_err("missing nested storage id")
                .missing_required_field(),
            Some("TSP.Reference.identifier")
        );
    }

    #[test]
    fn duplicate_selected_and_nested_fields_fail_before_buffa_last_wins() {
        let mut duplicate_note = slide(None, Some(42)).encode_to_vec();
        duplicate_note.extend_from_slice(&[0xda, 0x01, 0x02, 0x08, 0x2b]);
        assert_eq!(
            decode_slide(&duplicate_note)
                .expect_err("duplicate slide note")
                .duplicate_singular_field(),
            Some("KN.SlideArchive.note")
        );

        let mut native_with_title = slide(None, None);
        native_with_title.title_placeholder = Some(reference(51));
        let mut duplicate_title_source = native_with_title.encode_to_vec();
        duplicate_title_source.extend_from_slice(&[0x2a, 0x02, 0x08, 0x34]);
        assert_eq!(
            decode_slide(&duplicate_title_source)
                .expect_err("duplicate title placeholder")
                .duplicate_singular_field(),
            Some("KN.SlideArchive.titlePlaceholder")
        );

        let duplicate_identifier = [0x0a, 0x04, 0x08, 0x01, 0x08, 0x02];
        assert_eq!(
            decode_note(&duplicate_identifier)
                .expect_err("duplicate nested identifier")
                .duplicate_singular_field(),
            Some("TSP.Reference.identifier")
        );
    }

    #[test]
    fn noncanonical_and_invalid_utf8_fields_are_rejected() {
        let mut overlong_unknown = slide(None, None).encode_to_vec();
        overlong_unknown.extend_from_slice(&[0xa0, 0x00, 0x01]);
        assert_eq!(
            decode_slide(&overlong_unknown)
                .expect_err("overlong unknown key")
                .noncanonical_reason(),
            Some("protobuf field key")
        );

        let mut invalid_name = slide(None, None).encode_to_vec();
        invalid_name.extend_from_slice(&[0x52, 0x01, 0xff]);
        assert_eq!(
            decode_slide(&invalid_name)
                .expect_err("invalid UTF-8 name")
                .invalid_utf8_field(),
            Some("KN.SlideArchive.name")
        );

        let bad_bool = [
            0x0a, 0x02, 0x08, 0x01, 0x22, 0x02, 0x12, 0x00, 0x98, 0x01, 0x02,
        ];
        assert_eq!(
            decode_slide(&bad_bool)
                .expect_err("non-Boolean inDocument")
                .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
    }

    #[test]
    fn unknown_fields_remain_opaque_but_are_strictly_framed()
    -> Result<(), Box<dyn std::error::Error>> {
        let mut source = slide(Some("Named"), Some(42)).encode_to_vec();
        source.extend_from_slice(&[0x82, 0x02, 0x02, 0xff, 0x00]);
        let owner = decode_slide(&source)?;
        assert_eq!(owner.name(), Some("Named"));
        assert_eq!(owner.note().map(|value| value.identifier().get()), Some(42));
        Ok(())
    }

    #[test]
    fn selected_nested_unknown_fields_remain_opaque_when_lazy_views_are_forced()
    -> Result<(), Box<dyn std::error::Error>> {
        // The title/body references each carry a future unknown scalar. Their
        // lazy views are forced during the projection cross-check, so this
        // covers opaque nested bytes rather than only unknown slide fields.
        let source = [
            0x0a, 0x02, 0x08, 0x0b, // style = reference(11)
            0x22, 0x02, 0x12, 0x00, // transition.attributes = {}
            0x2a, 0x05, 0x08, 0x33, 0xa0, 0x06, 0x01, // title = reference(51), unknown 100
            0x32, 0x05, 0x08, 0x34, 0xa8, 0x06, 0x01, // body = reference(52), unknown 101
            0x98, 0x01, 0x01, // in_document = true
        ];
        let owner = decode_slide(&source)?;
        assert_eq!(
            owner
                .title_placeholder()
                .map(|reference| reference.identifier().get()),
            Some(51)
        );
        assert_eq!(
            owner
                .body_placeholder()
                .map(|reference| reference.identifier().get()),
            Some(52)
        );
        Ok(())
    }

    #[test]
    fn exact_resource_limits_pass_and_one_less_fails() -> Result<(), Box<dyn std::error::Error>> {
        let mut native_slide = slide(None, Some(42));
        native_slide.title_placeholder = Some(reference(51));
        native_slide.body_placeholder = Some(reference(52));
        let source = native_slide.encode_to_vec();
        let generous = options(&source, 3);
        let expected = decode_slide_notes_owner(&source, generous)?;

        let bytes_error = decode_slide_notes_owner(
            &source,
            DecodeOptions::new(source.len() - 1, source.len(), source.len() * 8, 3),
        );
        assert_eq!(
            bytes_error
                .expect_err("message bytes limit")
                .wire_resource_limit(),
            Some(WireResourceLimit::Bytes {
                observed: Some(source.len()),
                maximum: Some(source.len() - 1),
            })
        );

        let field_error = decode_slide_notes_owner(
            &source,
            DecodeOptions::new(source.len(), 1, source.len() * 8, 3),
        )
        .expect_err("field limit");
        assert_eq!(field_error.field_limit_values(), Some((2, 1)));

        let exact_work = strict_work_for_slide(&source)?;
        assert_eq!(
            decode_slide_notes_owner(
                &source,
                DecodeOptions::new(source.len(), source.len(), exact_work, 3),
            )?,
            expected
        );
        assert_eq!(
            decode_slide_notes_owner(
                &source,
                DecodeOptions::new(source.len(), source.len(), exact_work - 1, 3),
            )
            .expect_err("work limit")
            .work_limit_values(),
            Some((exact_work, exact_work - 1))
        );
        let nesting_error = decode_slide_notes_owner(
            &source,
            DecodeOptions::new(source.len(), source.len(), source.len() * 8, 1),
        );
        assert_eq!(
            nesting_error
                .expect_err("recursion limit")
                .wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: None,
                maximum: Some(1),
            })
        );
        Ok(())
    }

    #[test]
    fn invalid_resource_policy_has_content_free_limit_classification() {
        let source = slide(None, None).encode_to_vec();
        let error = decode_slide_notes_owner(
            &source,
            DecodeOptions::new(source.len(), source.len(), source.len() * 8, 0),
        )
        .expect_err("zero recursion limit");
        assert_eq!(
            error.wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: Some(0),
                maximum: Some(64),
            })
        );
    }

    fn strict_work_for_slide(source: &[u8]) -> Result<usize, Box<dyn std::error::Error>> {
        let native = kn::SlideArchive::decode(source)?;
        let style = native.style.encode_to_vec();
        let transition = native.transition.encode_to_vec();
        let attributes = native.transition.attributes.encode_to_vec();
        let title_placeholder = native
            .title_placeholder
            .map(|reference| reference.encode_to_vec());
        let body_placeholder = native
            .body_placeholder
            .map(|reference| reference.encode_to_vec());
        let note = native.note.map(|reference| reference.encode_to_vec());
        Ok(source
            .len()
            .checked_add(style.len())
            .and_then(|value| value.checked_add(transition.len()))
            .and_then(|value| value.checked_add(attributes.len()))
            .and_then(|value| value.checked_add(title_placeholder.as_ref().map_or(0, Vec::len)))
            .and_then(|value| value.checked_add(body_placeholder.as_ref().map_or(0, Vec::len)))
            .and_then(|value| value.checked_add(note.as_ref().map_or(0, Vec::len)))
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
            &[0x22, 0x01, 0x12],
            &[0x98, 0x01, 0x80],
        ];
        for source in malformed {
            assert!(decode_slide(source).is_err());
            assert!(decode_note(source).is_err());
        }
    }
}
