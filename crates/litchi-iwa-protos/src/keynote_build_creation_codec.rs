//! Bounded Buffa authoring for fresh Keynote audio/movie-start builds.
//!
//! The native build and chunk archives contain many fields that are not part
//! of the audio/movie-start contract.  This module owns the small canonical
//! writer used by the focused Keynote package adapter. It emits typed Buffa
//! views; it never constructs the generated native Prost messages and it
//! never decodes an existing build to manufacture a replacement.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The typed request, bounded writer, and wire oracle are kept together."
)]

use std::fmt;

use buffa::ViewEncode;

use crate::buffa_keynote_build_creation_generated::LitchiIwaProjection as projection;

const MAX_DEFAULT_OUTPUT_BYTES: usize = 16 * 1024;
const MAX_DEFAULT_WORK_BYTES: usize = 64 * 1024;
const MAX_DEFAULT_FIELDS: usize = 64;
// A build view owns three nested message views and a chunk view owns four;
// each operation also reserves its output vector.  Keep the default large
// enough for either canonical payload while still making the allocation
// accounting explicit to callers that provide a tighter policy.
const MAX_DEFAULT_ALLOCATIONS: usize = 5;

const BUILD_FIELDS: usize = 15;
const CHUNK_FIELDS: usize = 14;
const BUILD_VIEW_ALLOCATIONS: usize = 3;
const CHUNK_VIEW_ALLOCATIONS: usize = 4;

/// The identifiers and UUID halves needed to author one audio/movie-start pair.
///
/// The object identifiers are used only for archive references.  The UUID is
/// copied into both native chunk UUID locations, matching Keynote's writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct StartBuildWrite {
    drawable_identifier: u64,
    build_identifier: u64,
    chunk_identifier: u64,
    uuid_lower: u64,
    uuid_upper: u64,
    random_number_seed: u32,
}

impl StartBuildWrite {
    const fn new(
        drawable_identifier: u64,
        build_identifier: u64,
        chunk_identifier: u64,
        uuid_lower: u64,
        uuid_upper: u64,
        random_number_seed: u32,
    ) -> Self {
        Self {
            drawable_identifier,
            build_identifier,
            chunk_identifier,
            uuid_lower,
            uuid_upper,
            random_number_seed,
        }
    }
}

/// The identifiers and UUID halves needed to author one audio-start pair.
///
/// This low-level value is consumed only by the package adapter's bounded
/// writer. The semantic Keynote API allocates these identities privately.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StartAudioBuildWrite {
    inner: StartBuildWrite,
}

impl StartAudioBuildWrite {
    /// Construct a typed audio-start build request.
    #[must_use]
    pub const fn new(
        drawable_identifier: u64,
        build_identifier: u64,
        chunk_identifier: u64,
        uuid_lower: u64,
        uuid_upper: u64,
        random_number_seed: u32,
    ) -> Self {
        Self {
            inner: StartBuildWrite::new(
                drawable_identifier,
                build_identifier,
                chunk_identifier,
                uuid_lower,
                uuid_upper,
                random_number_seed,
            ),
        }
    }

    /// Return the drawable object identifier.
    #[must_use]
    pub const fn drawable_identifier(self) -> u64 {
        self.inner.drawable_identifier
    }

    /// Return the build object identifier.
    #[must_use]
    pub const fn build_identifier(self) -> u64 {
        self.inner.build_identifier
    }

    /// Return the build-chunk object identifier.
    #[must_use]
    pub const fn chunk_identifier(self) -> u64 {
        self.inner.chunk_identifier
    }

    /// Return the lower UUID half.
    #[must_use]
    pub const fn uuid_lower(self) -> u64 {
        self.inner.uuid_lower
    }

    /// Return the upper UUID half.
    #[must_use]
    pub const fn uuid_upper(self) -> u64 {
        self.inner.uuid_upper
    }

    /// Return the random seed copied into the animation attributes.
    #[must_use]
    pub const fn random_number_seed(self) -> u32 {
        self.inner.random_number_seed
    }
}

/// The identifiers and UUID halves needed to author one movie-start pair.
///
/// The semantic movie API does not expose this value. It is the typed seam
/// between its private identity allocator and the neutral Buffa writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StartMovieBuildWrite {
    inner: StartBuildWrite,
}

impl StartMovieBuildWrite {
    /// Construct a typed movie-start build request.
    #[must_use]
    pub const fn new(
        drawable_identifier: u64,
        build_identifier: u64,
        chunk_identifier: u64,
        uuid_lower: u64,
        uuid_upper: u64,
        random_number_seed: u32,
    ) -> Self {
        Self {
            inner: StartBuildWrite::new(
                drawable_identifier,
                build_identifier,
                chunk_identifier,
                uuid_lower,
                uuid_upper,
                random_number_seed,
            ),
        }
    }

    /// Return the drawable object identifier.
    #[must_use]
    pub const fn drawable_identifier(self) -> u64 {
        self.inner.drawable_identifier
    }

    /// Return the build object identifier.
    #[must_use]
    pub const fn build_identifier(self) -> u64 {
        self.inner.build_identifier
    }

    /// Return the build-chunk object identifier.
    #[must_use]
    pub const fn chunk_identifier(self) -> u64 {
        self.inner.chunk_identifier
    }

    /// Return the lower UUID half.
    #[must_use]
    pub const fn uuid_lower(self) -> u64 {
        self.inner.uuid_lower
    }

    /// Return the upper UUID half.
    #[must_use]
    pub const fn uuid_upper(self) -> u64 {
        self.inner.uuid_upper
    }

    /// Return the random seed copied into the animation attributes.
    #[must_use]
    pub const fn random_number_seed(self) -> u32 {
        self.inner.random_number_seed
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StartBuildKind {
    Audio,
    Movie,
}

impl StartBuildKind {
    const fn effect(self) -> &'static str {
        match self {
            Self::Audio => "apple:audio-start",
            Self::Movie => "apple:movie-start",
        }
    }
}

/// Finite resources for one typed build or chunk payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeOptions {
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_allocations: usize,
}

impl EncodeOptions {
    /// Construct an explicit finite output policy.
    #[must_use]
    pub const fn new(
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_allocations: usize,
    ) -> Self {
        Self {
            max_output_bytes,
            max_fields,
            max_work_bytes,
            max_allocations,
        }
    }

    /// Construct the conservative default policy for a fresh audio build payload.
    #[must_use]
    pub const fn for_write(_write: &StartAudioBuildWrite) -> Self {
        Self::new(
            MAX_DEFAULT_OUTPUT_BYTES,
            MAX_DEFAULT_FIELDS,
            MAX_DEFAULT_WORK_BYTES,
            MAX_DEFAULT_ALLOCATIONS,
        )
    }

    /// Construct the conservative default policy for a fresh movie build
    /// payload.
    #[must_use]
    pub const fn for_movie_write(_write: &StartMovieBuildWrite) -> Self {
        Self::new(
            MAX_DEFAULT_OUTPUT_BYTES,
            MAX_DEFAULT_FIELDS,
            MAX_DEFAULT_WORK_BYTES,
            MAX_DEFAULT_ALLOCATIONS,
        )
    }

    /// Replace the output-byte ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the encoded-field ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the aggregate measurement and encoding work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Replace the top-level allocation ceiling.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }
}

/// Exact finite resource consumption for one generated payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeReport {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    allocations: usize,
}

impl EncodeReport {
    /// Return the encoded payload size.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Return the number of emitted protobuf fields, including nested fields.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Return strict measurement plus encoding work.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Return the top-level allocation count.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

/// Encoded build/chunk payload and finite resource evidence.
#[derive(Debug, PartialEq, Eq)]
pub struct EncodeOutput {
    bytes: Vec<u8>,
    report: EncodeReport,
}

impl EncodeOutput {
    /// Borrow the encoded payload.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Return the exact resource report.
    #[must_use]
    pub const fn report(&self) -> EncodeReport {
        self.report
    }

    /// Consume the output and return its bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

/// A rejected typed authoring value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum InvalidInput {
    /// A required archive-object identifier was zero.
    ZeroIdentifier(&'static str),
    /// The build and chunk object identifiers were equal.
    DuplicateObjectIdentifier,
    /// Both UUID halves were zero.
    ZeroUuid,
}

/// A finite build-encoding resource ceiling.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncodeLimit {
    /// The generated payload exceeded the output ceiling.
    OutputBytes { observed: usize, maximum: usize },
    /// The generated payload exceeded the field ceiling.
    Fields { observed: usize, maximum: usize },
    /// Measurement plus encoding exceeded the work ceiling.
    WorkBytes { observed: usize, maximum: usize },
    /// The operation exceeded its allocation ceiling.
    Allocations { observed: usize, maximum: usize },
}

/// Failure from the bounded canonical build/chunk writer.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncodeError {
    /// A required identifier or UUID was invalid.
    InvalidInput(InvalidInput),
    /// A finite caller budget was exceeded.
    Resource(EncodeLimit),
    /// The destination vector could not be reserved.
    Allocation { amount: usize },
    /// Buffa rejected the generated view or bound.
    Buffa(buffa::EncodeError),
    /// The measured and emitted lengths differed.
    Verification,
}

impl fmt::Display for EncodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidInput(input) => {
                write!(formatter, "invalid Keynote build input: {input:?}")
            },
            Self::Resource(_) => formatter.write_str("Keynote build encoding limit exceeded"),
            Self::Allocation { amount } => {
                write!(
                    formatter,
                    "Keynote build output allocation failed ({amount} bytes)"
                )
            },
            Self::Buffa(error) => error.fmt(formatter),
            Self::Verification => formatter.write_str("Keynote build encoding verification failed"),
        }
    }
}

impl std::error::Error for EncodeError {}

impl From<buffa::EncodeError> for EncodeError {
    fn from(error: buffa::EncodeError) -> Self {
        Self::Buffa(error)
    }
}

/// Encode the canonical `KN.BuildArchive` audio-start payload.
pub fn encode_start_audio_build(
    write: &StartAudioBuildWrite,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    encode_start_build(&write.inner, StartBuildKind::Audio, options)
}

/// Encode the canonical `KN.BuildArchive` movie-start payload.
pub fn encode_start_movie_build(
    write: &StartMovieBuildWrite,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    encode_start_build(&write.inner, StartBuildKind::Movie, options)
}

/// Encode the canonical `KN.BuildChunkArchive` audio-start payload.
pub fn encode_start_audio_chunk(
    write: &StartAudioBuildWrite,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    encode_start_chunk(&write.inner, options)
}

/// Encode the canonical `KN.BuildChunkArchive` movie-start payload.
pub fn encode_start_movie_chunk(
    write: &StartMovieBuildWrite,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    encode_start_chunk(&write.inner, options)
}

fn encode_start_build(
    write: &StartBuildWrite,
    kind: StartBuildKind,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    validate_write(write)?;
    let report = encode_report(
        build_encoded_len(write, kind)?,
        BUILD_FIELDS,
        BUILD_VIEW_ALLOCATIONS,
    )?;
    preflight(report, options)?;
    let view = build_view(write, kind);
    encode_view(&view, report, options)
}

fn encode_start_chunk(
    write: &StartBuildWrite,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    validate_write(write)?;
    let report = encode_report(
        chunk_encoded_len(write)?,
        CHUNK_FIELDS,
        CHUNK_VIEW_ALLOCATIONS,
    )?;
    preflight(report, options)?;
    let view = chunk_view(write);
    encode_view(&view, report, options)
}

fn validate_write(write: &StartBuildWrite) -> Result<(), EncodeError> {
    for (identifier, name) in [
        (write.drawable_identifier, "drawable_identifier"),
        (write.build_identifier, "build_identifier"),
        (write.chunk_identifier, "chunk_identifier"),
    ] {
        if identifier == 0 {
            return Err(EncodeError::InvalidInput(InvalidInput::ZeroIdentifier(
                name,
            )));
        }
    }
    if write.build_identifier == write.chunk_identifier {
        return Err(EncodeError::InvalidInput(
            InvalidInput::DuplicateObjectIdentifier,
        ));
    }
    if write.uuid_lower == 0 && write.uuid_upper == 0 {
        return Err(EncodeError::InvalidInput(InvalidInput::ZeroUuid));
    }
    Ok(())
}

fn encode_view<'a>(
    view: &impl ViewEncode<'a>,
    report: EncodeReport,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    // The static wire-length oracle and preflight run before this view exists,
    // so MessageFieldView::set cannot allocate nested boxes for a rejected
    // request.  Measuring the generated view here is only a parity check.
    let measured_output_bytes =
        usize::try_from(view.try_encoded_len()?).map_err(|_| EncodeError::Verification)?;
    if measured_output_bytes != report.output_bytes {
        return Err(EncodeError::Verification);
    }

    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(report.output_bytes)
        .map_err(|_| EncodeError::Allocation {
            amount: report.output_bytes,
        })?;
    if bytes.capacity() < report.output_bytes {
        return Err(EncodeError::Allocation {
            amount: report.output_bytes,
        });
    }
    let maximum = u32::try_from(options.max_output_bytes).unwrap_or(u32::MAX);
    let encoded = view.try_encode_bounded(maximum, &mut bytes)?;
    if usize::try_from(encoded).ok() != Some(report.output_bytes)
        || bytes.len() != report.output_bytes
    {
        return Err(EncodeError::Verification);
    }
    Ok(EncodeOutput { bytes, report })
}

fn encode_report(
    output_bytes: usize,
    fields: usize,
    view_allocations: usize,
) -> Result<EncodeReport, EncodeError> {
    let allocations = view_allocations
        .checked_add(1)
        .ok_or(EncodeError::Verification)?;
    let work_bytes = output_bytes
        .checked_mul(3)
        .and_then(|work| work.checked_add(fields))
        .ok_or(EncodeError::Verification)?;
    Ok(EncodeReport {
        output_bytes,
        fields,
        work_bytes,
        allocations,
    })
}

fn preflight(report: EncodeReport, options: EncodeOptions) -> Result<(), EncodeError> {
    let limit = if report.output_bytes > options.max_output_bytes {
        Some(EncodeLimit::OutputBytes {
            observed: report.output_bytes,
            maximum: options.max_output_bytes,
        })
    } else if report.fields > options.max_fields {
        Some(EncodeLimit::Fields {
            observed: report.fields,
            maximum: options.max_fields,
        })
    } else if report.work_bytes > options.max_work_bytes {
        Some(EncodeLimit::WorkBytes {
            observed: report.work_bytes,
            maximum: options.max_work_bytes,
        })
    } else if report.allocations > options.max_allocations {
        Some(EncodeLimit::Allocations {
            observed: report.allocations,
            maximum: options.max_allocations,
        })
    } else {
        None
    };
    if let Some(limit) = limit {
        return Err(EncodeError::Resource(limit));
    }
    Ok(())
}

fn build_encoded_len(write: &StartBuildWrite, kind: StartBuildKind) -> Result<usize, EncodeError> {
    let drawable = required_len(message_field_len(
        1,
        varint_field_len(1, write.drawable_identifier),
    ))?;
    let mut animation_attributes = required_len(string_field_len(1, b"In"))?;
    animation_attributes = add_len(
        animation_attributes,
        required_len(string_field_len(2, kind.effect().as_bytes()))?,
    )?;
    animation_attributes = add_len(animation_attributes, fixed64_field_len(3))?;
    animation_attributes = add_len(animation_attributes, fixed64_field_len(5))?;
    animation_attributes = add_len(
        animation_attributes,
        varint_field_len(11, u64::from(write.random_number_seed)),
    )?;
    animation_attributes = add_len(animation_attributes, bool_field_len(16))?;

    let mut attributes = varint_field_len(4, 1);
    attributes = add_len(attributes, fixed64_field_len(17))?;
    attributes = add_len(
        attributes,
        required_len(message_field_len(18, animation_attributes))?,
    )?;

    let mut total = drawable;
    total = add_len(total, required_len(string_field_len(2, b"All at Once"))?)?;
    total = add_len(total, fixed64_field_len(3))?;
    total = add_len(total, required_len(message_field_len(4, attributes))?)?;
    add_len(total, varint_field_len(5, 1))
}

fn chunk_encoded_len(write: &StartBuildWrite) -> Result<usize, EncodeError> {
    let build = required_len(message_field_len(
        1,
        varint_field_len(1, write.build_identifier),
    ))?;
    let uuid = add_len(
        varint_field_len(1, write.uuid_lower),
        varint_field_len(2, write.uuid_upper),
    )?;
    let mut build_chunk_identifier = required_len(message_field_len(1, uuid))?;
    build_chunk_identifier = add_len(build_chunk_identifier, varint_field_len(2, 1))?;

    let mut total = build;
    total = add_len(total, fixed64_field_len(3))?;
    total = add_len(total, fixed64_field_len(4))?;
    total = add_len(total, bool_field_len(5))?;
    total = add_len(total, bool_field_len(6))?;
    total = add_len(
        total,
        required_len(message_field_len(7, build_chunk_identifier))?,
    )?;
    add_len(total, required_len(message_field_len(8, uuid))?)
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn key_len(field_number: u32, wire_type: u8) -> usize {
    varint_len((u64::from(field_number) << 3) | u64::from(wire_type))
}

fn varint_field_len(field_number: u32, value: u64) -> usize {
    key_len(field_number, 0) + varint_len(value)
}

fn bool_field_len(field_number: u32) -> usize {
    varint_field_len(field_number, 0)
}

fn fixed64_field_len(field_number: u32) -> usize {
    key_len(field_number, 1) + 8
}

fn string_field_len(field_number: u32, value: &[u8]) -> Option<usize> {
    message_field_len(field_number, value.len())
}

fn message_field_len(field_number: u32, payload_len: usize) -> Option<usize> {
    let payload_len = u64::try_from(payload_len).ok()?;
    key_len(field_number, 2)
        .checked_add(varint_len(payload_len))?
        .checked_add(usize::try_from(payload_len).ok()?)
}

fn required_len(length: Option<usize>) -> Result<usize, EncodeError> {
    length.ok_or(EncodeError::Verification)
}

fn add_len(left: usize, right: usize) -> Result<usize, EncodeError> {
    left.checked_add(right).ok_or(EncodeError::Verification)
}

fn reference(identifier: u64) -> projection::ReferenceView<'static> {
    projection::ReferenceView {
        identifier,
        ..Default::default()
    }
}

fn uuid(lower: u64, upper: u64) -> projection::UUIDView<'static> {
    projection::UUIDView {
        lower,
        upper,
        ..Default::default()
    }
}

fn build_view(
    write: &StartBuildWrite,
    kind: StartBuildKind,
) -> projection::BuildArchiveView<'static> {
    let animation_attributes = projection::AnimationAttributesArchiveView {
        animation_type: Some("In"),
        effect: Some(kind.effect()),
        duration: Some(0.5),
        delay: Some(0.0),
        random_number_seed: Some(write.random_number_seed),
        writing_direction_is_rtl: Some(false),
    };
    let attributes = projection::BuildAttributesArchiveView {
        event_trigger: Some(1),
        chart_rotation_3_d: Some(60.0),
        animation_attributes: buffa::MessageFieldView::set(animation_attributes),
    };
    projection::BuildArchiveView {
        drawable: buffa::MessageFieldView::set(reference(write.drawable_identifier)),
        delivery: "All at Once",
        duration: Some(0.0),
        attributes: buffa::MessageFieldView::set(attributes),
        chunk_id_seed: Some(1),
        __buffa_required_seen_0: 0,
    }
}

fn chunk_view(write: &StartBuildWrite) -> projection::BuildChunkArchiveView<'static> {
    let identifier = projection::BuildChunkIdentifierArchiveView {
        build_id: buffa::MessageFieldView::set(uuid(write.uuid_lower, write.uuid_upper)),
        build_chunk_id: Some(1),
    };
    projection::BuildChunkArchiveView {
        build: buffa::MessageFieldView::set(reference(write.build_identifier)),
        delay: Some(0.0),
        duration: Some(0.5),
        automatic: Some(false),
        referent: Some(true),
        build_chunk_identifier: buffa::MessageFieldView::set(identifier),
        build_id: buffa::MessageFieldView::set(uuid(write.uuid_lower, write.uuid_upper)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use prost::Message as _;

    fn write() -> StartAudioBuildWrite {
        StartAudioBuildWrite::new(11, 17, 19, 0x0102_0304_0506_0708, 0x1112_1314_1516_1718, 23)
    }

    fn movie_write() -> StartMovieBuildWrite {
        StartMovieBuildWrite::new(11, 17, 19, 0x0102_0304_0506_0708, 0x1112_1314_1516_1718, 23)
    }

    #[allow(deprecated)]
    #[test]
    fn build_matches_native_wire_shape() -> Result<(), Box<dyn std::error::Error>> {
        let input = write();
        let output = encode_start_audio_build(&input, EncodeOptions::for_write(&input))?;
        let expected = crate::kn::BuildArchive {
            drawable: Some(crate::tsp::Reference {
                identifier: input.drawable_identifier(),
                ..Default::default()
            }),
            delivery: "All at Once".to_owned(),
            duration: Some(0.0),
            attributes: crate::kn::BuildAttributesArchive {
                animation_attributes: Some(crate::kn::AnimationAttributesArchive {
                    animation_type: Some("In".to_owned()),
                    effect: Some("apple:audio-start".to_owned()),
                    duration: Some(0.5),
                    delay: Some(0.0),
                    random_number_seed: Some(input.random_number_seed()),
                    writing_direction_is_rtl: Some(false),
                    ..Default::default()
                }),
                event_trigger: Some(1),
                chart_rotation3_d: Some(60.0),
                ..Default::default()
            },
            chunk_id_seed: Some(1),
        };
        assert_eq!(output.bytes(), expected.encode_to_vec());
        Ok(())
    }

    #[test]
    fn chunk_matches_native_wire_shape() -> Result<(), Box<dyn std::error::Error>> {
        let input = write();
        let output = encode_start_audio_chunk(&input, EncodeOptions::for_write(&input))?;
        let uuid = crate::tsp::Uuid {
            lower: input.uuid_lower(),
            upper: input.uuid_upper(),
        };
        let expected = crate::kn::BuildChunkArchive {
            build: Some(crate::tsp::Reference {
                identifier: input.build_identifier(),
                ..Default::default()
            }),
            delay: Some(0.0),
            duration: Some(0.5),
            automatic: Some(false),
            referent: Some(true),
            build_chunk_identifier: Some(crate::kn::BuildChunkIdentifierArchive {
                build_id: Some(uuid),
                build_chunk_id: Some(1),
            }),
            build_id: Some(uuid),
            ..Default::default()
        };
        assert_eq!(output.bytes(), expected.encode_to_vec());
        Ok(())
    }

    #[allow(deprecated)]
    #[test]
    fn movie_build_matches_native_wire_shape() -> Result<(), Box<dyn std::error::Error>> {
        let input = movie_write();
        let output = encode_start_movie_build(&input, EncodeOptions::for_movie_write(&input))?;
        let expected = crate::kn::BuildArchive {
            drawable: Some(crate::tsp::Reference {
                identifier: input.drawable_identifier(),
                ..Default::default()
            }),
            delivery: "All at Once".to_owned(),
            duration: Some(0.0),
            attributes: crate::kn::BuildAttributesArchive {
                animation_attributes: Some(crate::kn::AnimationAttributesArchive {
                    animation_type: Some("In".to_owned()),
                    effect: Some("apple:movie-start".to_owned()),
                    duration: Some(0.5),
                    delay: Some(0.0),
                    random_number_seed: Some(input.random_number_seed()),
                    writing_direction_is_rtl: Some(false),
                    ..Default::default()
                }),
                event_trigger: Some(1),
                chart_rotation3_d: Some(60.0),
                ..Default::default()
            },
            chunk_id_seed: Some(1),
        };
        assert_eq!(output.bytes(), expected.encode_to_vec());
        Ok(())
    }

    #[test]
    fn movie_chunk_shares_the_audio_wire_shape() -> Result<(), Box<dyn std::error::Error>> {
        let audio = write();
        let movie = movie_write();
        let audio_output = encode_start_audio_chunk(&audio, EncodeOptions::for_write(&audio))?;
        let movie_output =
            encode_start_movie_chunk(&movie, EncodeOptions::for_movie_write(&movie))?;
        assert_eq!(movie_output.bytes(), audio_output.bytes());
        assert_eq!(movie_output.report(), audio_output.report());
        Ok(())
    }

    #[test]
    fn movie_build_enforces_the_same_finite_field_budget() {
        let input = movie_write();
        let options = EncodeOptions::for_movie_write(&input).with_max_fields(BUILD_FIELDS - 1);
        let error = encode_start_movie_build(&input, options).expect_err("field budget");
        assert_eq!(
            error,
            EncodeError::Resource(EncodeLimit::Fields {
                observed: BUILD_FIELDS,
                maximum: BUILD_FIELDS - 1,
            })
        );
    }

    #[test]
    fn movie_and_audio_builds_differ_only_by_the_typed_effect() {
        let audio = write();
        let movie = movie_write();
        let audio = crate::kn::BuildArchive::decode(
            encode_start_audio_build(&audio, EncodeOptions::for_write(&audio))
                .expect("audio build")
                .bytes(),
        )
        .expect("decode audio build");
        let movie = crate::kn::BuildArchive::decode(
            encode_start_movie_build(&movie, EncodeOptions::for_movie_write(&movie))
                .expect("movie build")
                .bytes(),
        )
        .expect("decode movie build");
        let audio_attributes = audio
            .attributes
            .animation_attributes
            .expect("audio animation attributes");
        let movie_attributes = movie
            .attributes
            .animation_attributes
            .expect("movie animation attributes");
        assert_eq!(audio.drawable, movie.drawable);
        assert_eq!(audio.delivery, movie.delivery);
        assert_eq!(audio.duration, movie.duration);
        assert_eq!(
            audio.attributes.event_trigger,
            movie.attributes.event_trigger
        );
        assert_eq!(
            audio.attributes.chart_rotation3_d,
            movie.attributes.chart_rotation3_d
        );
        assert_eq!(audio.chunk_id_seed, movie.chunk_id_seed);
        assert_eq!(
            audio_attributes.animation_type,
            movie_attributes.animation_type
        );
        assert_eq!(audio_attributes.duration, movie_attributes.duration);
        assert_eq!(audio_attributes.delay, movie_attributes.delay);
        assert_eq!(
            audio_attributes.random_number_seed,
            movie_attributes.random_number_seed
        );
        assert_eq!(
            audio_attributes.writing_direction_is_rtl,
            movie_attributes.writing_direction_is_rtl
        );
        assert_eq!(
            audio_attributes.effect.as_deref(),
            Some("apple:audio-start")
        );
        assert_eq!(
            movie_attributes.effect.as_deref(),
            Some("apple:movie-start")
        );
    }

    #[test]
    fn invalid_object_graph_is_rejected_before_encoding() {
        let input = StartAudioBuildWrite::new(1, 2, 2, 3, 4, 5);
        let error = encode_start_audio_build(&input, EncodeOptions::for_write(&input))
            .expect_err("duplicate object identifiers must be rejected");
        assert_eq!(
            error,
            EncodeError::InvalidInput(InvalidInput::DuplicateObjectIdentifier)
        );
    }

    #[test]
    fn finite_field_budget_is_checked_before_allocation() {
        let input = write();
        let options = EncodeOptions::for_write(&input).with_max_fields(BUILD_FIELDS - 1);
        let error = encode_start_audio_build(&input, options).expect_err("field budget");
        assert_eq!(
            error,
            EncodeError::Resource(EncodeLimit::Fields {
                observed: BUILD_FIELDS,
                maximum: BUILD_FIELDS - 1,
            })
        );
    }

    #[test]
    fn nested_view_allocation_budget_is_checked_before_view_construction() {
        let input = write();
        let options = EncodeOptions::for_write(&input).with_max_allocations(3);
        let error = encode_start_audio_build(&input, options).expect_err("allocation budget");
        assert_eq!(
            error,
            EncodeError::Resource(EncodeLimit::Allocations {
                observed: BUILD_VIEW_ALLOCATIONS + 1,
                maximum: 3,
            })
        );
    }
}
