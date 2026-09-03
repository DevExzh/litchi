//! Strict private Buffa projection for one Pages media discriminator.
//!
//! Pages movies and audio clips share `TSD.MovieArchive`. The graph reader
//! needs only the `audioOnly` and `is_live_video` discriminators while discovering body
//! attachments; the complete archive remains caller-owned and is decoded by
//! the existing graph validation path. A handwritten wire pass owns framing,
//! canonical scalar validation, and resource accounting before a private
//! Buffa lazy view is forced. Unknown source fields are never materialized or rewritten here.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict semantic preflight intentionally precedes the low-level wire reader."
)]

use std::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_pages_media_generated::LitchiIwaProjection as projection;

const AUDIO_ONLY_FIELD: u32 = 9;
const IS_LIVE_VIDEO_FIELD: u32 = 30;
const MAX_RECURSION_LIMIT: u32 = 64;

/// Explicit finite resource policy for one Pages media payload.
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

    /// Build a conservative policy from one known source length.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(4).max(1),
            bytes.saturating_mul(8).max(1),
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
}

/// Borrow-free selected media discriminator.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MovieAudioFlagSnapshot {
    audio_only: Option<bool>,
    is_live_video: Option<bool>,
}

impl MovieAudioFlagSnapshot {
    /// Native `audioOnly` field, preserving field presence.
    #[must_use]
    pub const fn audio_only(self) -> Option<bool> {
        self.audio_only
    }

    /// Native `is_live_video` field, preserving field presence.
    #[must_use]
    pub const fn is_live_video(self) -> Option<bool> {
        self.is_live_video
    }
}

/// Selected `TSD.MovieArchive` media flags used by Pages discovery.
///
/// The historical name [`MovieAudioFlagSnapshot`] remains the concrete type
/// so downstream internal callers compiled against the first projection keep
/// their source compatibility. This alias describes the complete selected
/// flag set without exposing the native archive or raw IDs.
pub type MovieMediaFlagSnapshot = MovieAudioFlagSnapshot;

/// Failure from strict Pages media preflight or its Buffa cross-check.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    MessageByteLimit { observed: usize, maximum: usize },
    RecursionLimit { observed: u32, maximum: u32 },
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
    Projection,
}

impl DecodeError {
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

    /// Return the duplicated singular field, when applicable.
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
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::MessageByteLimit { observed, maximum } => write!(
                formatter,
                "Pages media projection observed {observed} message bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::RecursionLimit { observed, maximum } => write!(
                formatter,
                "Pages media projection reached recursion depth {observed}; maximum is {maximum}"
            ),
            DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Pages media projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Pages media projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Projection => formatter
                .write_str("Pages media strict preflight disagrees with the Buffa projection"),
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

/// Decode the selected `TSD.MovieArchive` media discriminators.
pub fn decode_movie_media_flags(
    source: &[u8],
    options: DecodeOptions,
) -> Result<MovieMediaFlagSnapshot, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight(source, options, &mut budget)?;
    let view: projection::MovieAudioFlagArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    let projected = MovieMediaFlagSnapshot {
        audio_only: view.audio_only,
        is_live_video: view.is_live_video,
    };
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

/// Decode the selected `TSD.MovieArchive` media discriminators.
///
/// This name is retained for compatibility with the original audio-only
/// projection. It now also validates and returns `is_live_video`, ensuring
/// movie and audio discovery share one strict wire contract.
pub fn decode_movie_audio_only(
    source: &[u8],
    options: DecodeOptions,
) -> Result<MovieMediaFlagSnapshot, DecodeError> {
    decode_movie_media_flags(source, options)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let max_buffa_message_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
    if options.max_message_bytes > max_buffa_message_bytes {
        return Err(DecodeError {
            kind: DecodeErrorKind::MessageByteLimit {
                observed: options.max_message_bytes,
                maximum: max_buffa_message_bytes,
            },
        });
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError {
            kind: DecodeErrorKind::MessageByteLimit {
                observed: source.len(),
                maximum: options.max_message_bytes,
            },
        });
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
        let observed = self.work_bytes.saturating_add(bytes.saturating_mul(2));
        if observed > self.max_work_bytes {
            return Err(DecodeError::work_limit(observed, self.max_work_bytes));
        }
        self.work_bytes = observed;
        Ok(())
    }
}

fn preflight(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<MovieAudioFlagSnapshot, DecodeError> {
    budget.charge_message(source.len())?;
    let mut audio_only = None;
    let mut is_live_video = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            AUDIO_ONLY_FIELD => {
                if audio_only.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSD.MovieArchive.audioOnly",
                    ));
                }
                audio_only = Some(require_canonical_bool(field.varint()?)?);
            },
            IS_LIVE_VIDEO_FIELD => {
                if is_live_video.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSD.MovieArchive.is_live_video",
                    ));
                }
                is_live_video = Some(require_canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    Ok(MovieAudioFlagSnapshot {
        audio_only,
        is_live_video,
    })
}

fn require_canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}

#[derive(Clone, Copy, Debug)]
enum StrictValue {
    Varint(u64),
    Fixed64,
    LengthDelimited,
    Group,
    Fixed32,
}

#[derive(Clone, Copy, Debug)]
struct StrictField {
    number: u32,
    wire_type: buffa::encoding::WireType,
    value: StrictValue,
}

impl StrictField {
    fn varint(self) -> Result<u64, DecodeError> {
        if self.wire_type != buffa::encoding::WireType::Varint {
            return Err(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected: buffa::encoding::WireType::Varint as u8,
                actual: self.wire_type as u8,
            }
            .into());
        }
        match self.value {
            StrictValue::Varint(value) => Ok(value),
            StrictValue::Fixed64 | StrictValue::LengthDelimited => Err(DecodeError::projection()),
            StrictValue::Group | StrictValue::Fixed32 => Err(DecodeError::projection()),
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum ParseItem {
    Field(StrictField),
    EndGroup(u32),
}

fn next_strict_field(
    source: &mut &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Option<StrictField>, DecodeError> {
    match parse_strict_field(source, options.recursion_limit, budget)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(number)) => {
            Err(buffa::DecodeError::InvalidEndGroup(number).into())
        },
        None => Ok(None),
    }
}

fn parse_strict_field(
    source: &mut &[u8],
    recursion_limit: u32,
    budget: &mut Budget,
) -> Result<Option<ParseItem>, DecodeError> {
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
            take_exact(source, length)?;
            StrictValue::LengthDelimited
        },
        buffa::encoding::WireType::StartGroup => {
            let child_limit = recursion_limit.checked_sub(1).ok_or_else(|| {
                // Report the remaining local budget rather than the global
                // policy ceiling. This keeps nested unknown-group failures
                // actionable when callers intentionally choose a shallow
                // limit, while Buffa receives the same strict bound below.
                DecodeError::recursion_limit(recursion_limit.saturating_add(1), recursion_limit)
            })?;
            skip_strict_group(source, field_number, child_limit, budget)?;
            StrictValue::Group
        },
        buffa::encoding::WireType::EndGroup => {
            return Ok(Some(ParseItem::EndGroup(field_number)));
        },
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
    reason = "Focused media-wire tests use explicit fixture expectations."
)]
mod tests {
    use prost::Message as _;

    use super::*;
    use crate::tsd;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
    }

    #[test]
    fn canonical_movie_audio_flag_matches_projection() {
        let source = tsd::MovieArchive {
            audio_only: Some(true),
            is_live_video: Some(true),
            ..tsd::MovieArchive::default()
        }
        .encode_to_vec();
        let snapshot =
            decode_movie_media_flags(&source, options(&source)).expect("canonical movie payload");
        assert_eq!(snapshot.audio_only(), Some(true));
        assert_eq!(snapshot.is_live_video(), Some(true));

        let compatibility = decode_movie_audio_only(&source, options(&source))
            .expect("compatibility movie payload");
        assert_eq!(compatibility, snapshot);
    }

    #[test]
    fn absence_and_false_presence_are_distinct() {
        let absent = tsd::MovieArchive::default().encode_to_vec();
        assert_eq!(
            decode_movie_audio_only(&absent, options(&absent))
                .expect("absent discriminator")
                .audio_only(),
            None
        );
        assert_eq!(
            decode_movie_media_flags(&absent, options(&absent))
                .expect("absent live-video discriminator")
                .is_live_video(),
            None
        );
        let explicit_false = tsd::MovieArchive {
            audio_only: Some(false),
            is_live_video: Some(false),
            ..tsd::MovieArchive::default()
        }
        .encode_to_vec();
        let snapshot = decode_movie_media_flags(&explicit_false, options(&explicit_false))
            .expect("explicit false discriminator");
        assert_eq!(snapshot.audio_only(), Some(false));
        assert_eq!(snapshot.is_live_video(), Some(false));
    }

    #[test]
    fn unknown_fields_are_ignored_without_reencoding() {
        let mut source = tsd::MovieArchive {
            audio_only: Some(true),
            is_live_video: Some(false),
            ..tsd::MovieArchive::default()
        }
        .encode_to_vec();
        source.extend([0x82, 0x06, 0x03, 0xaa, 0xbb, 0xcc]);
        assert_eq!(
            decode_movie_audio_only(&source, options(&source))
                .expect("unknown movie field")
                .audio_only(),
            Some(true)
        );
        assert_eq!(
            decode_movie_media_flags(&source, options(&source))
                .expect("unknown movie field")
                .is_live_video(),
            Some(false)
        );
    }

    #[test]
    fn duplicate_and_noncanonical_bools_are_rejected_before_projection() {
        let duplicate = [0x48, 0x01, 0x48, 0x00];
        let error = decode_movie_audio_only(&duplicate, options(&duplicate))
            .expect_err("duplicate audioOnly");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSD.MovieArchive.audioOnly")
        );

        let noncanonical = [0x48, 0x02];
        let error = decode_movie_audio_only(&noncanonical, options(&noncanonical))
            .expect_err("noncanonical bool");
        assert_eq!(
            error.noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );

        let duplicate_live_video = [0xf0, 0x01, 0x01, 0xf0, 0x01, 0x00];
        let error = decode_movie_media_flags(&duplicate_live_video, options(&duplicate_live_video))
            .expect_err("duplicate is_live_video");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSD.MovieArchive.is_live_video")
        );

        let noncanonical_live_video = [0xf0, 0x01, 0x02];
        let error =
            decode_movie_media_flags(&noncanonical_live_video, options(&noncanonical_live_video))
                .expect_err("noncanonical is_live_video");
        assert_eq!(
            error.noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
    }

    #[test]
    fn field_and_work_limits_are_finite() {
        let source = [0x10, 0x01, 0x18, 0x01, 0xf0, 0x01, 0x00];
        let field_limited = DecodeOptions::new(source.len(), 1, source.len() * 8, 8);
        let error = decode_movie_audio_only(&source, field_limited).expect_err("field limit");
        assert_eq!(error.field_limit_values(), Some((2, 1)));

        let work_limited = DecodeOptions::new(source.len(), 8, source.len(), 8);
        let error = decode_movie_audio_only(&source, work_limited).expect_err("work limit");
        assert_eq!(
            error.work_limit_values(),
            Some((source.len() * 2, source.len()))
        );
    }

    #[test]
    fn nested_unknown_group_reports_remaining_recursion_budget() {
        // Field 100 starts an unknown group; field 101 starts a second group
        // before the first one closes. A one-level policy permits the outer
        // group but rejects this nested start before Buffa is forced.
        let source = [0xa3, 0x06, 0xab, 0x06, 0xac, 0x06, 0xa4, 0x06];
        let error = decode_movie_audio_only(
            &source,
            DecodeOptions::new(source.len(), 16, source.len() * 8, 1),
        )
        .expect_err("nested unknown group must hit the local depth ceiling");
        assert_eq!(error.recursion_limit_values(), Some((1, 0)));
    }
}
