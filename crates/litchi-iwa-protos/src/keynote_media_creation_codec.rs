//! Bounded Buffa authoring for fresh Keynote audio/movie archives.
//!
//! A Keynote media object has a large generated `TSD.MovieArchive` type, but a
//! newly authored audio/movie object only needs a small, stable subset of its
//! fields.  This module keeps that authoring surface typed and format-neutral,
//! then encodes a private Buffa view directly.  The native generated Prost
//! graph is used only by the test-only differential oracle.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The public value types precede their bounded encoding helpers."
)]

use std::fmt;

use buffa::ViewEncode as _;

use crate::buffa_keynote_media_creation_generated::LitchiIwaProjection as buffa_tsd;

const AUDIO_REFERENCES: usize = 5;
const MOVIE_REFERENCES_WITHOUT_POSTER: usize = 5;
const MOVIE_REFERENCES_WITH_POSTER: usize = 6;
const MAX_DEFAULT_OUTPUT_BYTES: usize = 16 * 1024;
const MAX_DEFAULT_WORK_BYTES: usize = 64 * 1024;
const MAX_DEFAULT_FIELDS: usize = 128;
const MAX_DEFAULT_REFERENCES: usize = 8;
const MAX_DEFAULT_ALLOCATIONS: usize = 16;

/// A finite point in the media drawable's geometry.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Point {
    x: f32,
    y: f32,
}

impl Point {
    /// Construct a point.  Non-finite values are rejected when encoding.
    #[must_use]
    pub const fn new(x: f32, y: f32) -> Self {
        Self { x, y }
    }

    /// Return the horizontal coordinate.
    #[must_use]
    pub const fn x(self) -> f32 {
        self.x
    }

    /// Return the vertical coordinate.
    #[must_use]
    pub const fn y(self) -> f32 {
        self.y
    }
}

/// A finite media size.  Zero is accepted because an audio-only archive's
/// native `naturalSize` is canonically `0 × 0`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Size {
    width: f32,
    height: f32,
}

impl Size {
    /// Construct a size.  Negative or non-finite values are rejected when
    /// encoding.
    #[must_use]
    pub const fn new(width: f32, height: f32) -> Self {
        Self { width, height }
    }

    /// Return the width.
    #[must_use]
    pub const fn width(self) -> f32 {
        self.width
    }

    /// Return the height.
    #[must_use]
    pub const fn height(self) -> f32 {
        self.height
    }
}

/// The only drawable geometry values authored by the fresh media profile.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Geometry {
    position: Point,
    size: Size,
    flags: Option<u32>,
    angle: Option<f32>,
}

impl Geometry {
    /// Construct drawable geometry.  Scalar validity is checked when
    /// encoding so callers can build values before choosing a failure policy.
    #[must_use]
    pub const fn new(position: Point, size: Size, flags: Option<u32>, angle: Option<f32>) -> Self {
        Self {
            position,
            size,
            flags,
            angle,
        }
    }

    /// Return the position.
    #[must_use]
    pub const fn position(self) -> Point {
        self.position
    }

    /// Return the drawable size.
    #[must_use]
    pub const fn size(self) -> Size {
        self.size
    }

    /// Return native geometry flags.
    #[must_use]
    pub const fn flags(self) -> Option<u32> {
        self.flags
    }

    /// Return the clockwise angle in native floating-point units.
    #[must_use]
    pub const fn angle(self) -> Option<f32> {
        self.angle
    }
}

/// Media-specific values that distinguish an audio-only archive from a movie
/// archive while retaining the native defaults shared by both.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum MediaKind {
    /// An audio-only archive.  Native `naturalSize` is `0 × 0` and no poster
    /// data reference is emitted.
    Audio,
    /// A movie archive, with an optional poster data object and native movie
    /// size.  A poster identifier is omitted when the source has no poster.
    Movie {
        /// Optional poster data object identifier.
        poster_data_identifier: Option<u64>,
        /// Native original and natural movie size.
        natural_size: Size,
        /// Native alpha-capable poster flag.
        poster_image_generated_with_alpha_support: bool,
    },
}

impl MediaKind {
    /// Construct an audio-only kind.
    #[must_use]
    pub const fn audio() -> Self {
        Self::Audio
    }

    /// Construct a movie kind.
    #[must_use]
    pub const fn movie(
        poster_data_identifier: Option<u64>,
        natural_size: Size,
        poster_image_generated_with_alpha_support: bool,
    ) -> Self {
        Self::Movie {
            poster_data_identifier,
            natural_size,
            poster_image_generated_with_alpha_support,
        }
    }
}

/// Typed input for one newly authored Keynote audio/movie archive.
///
/// The archive's object/data identifiers are validated as non-zero at the
/// encoding boundary.  Keeping them as scalar values makes integration with
/// package allocators straightforward while preventing invalid references from
/// reaching the wire.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MediaArchiveWrite {
    parent_identifier: u64,
    style_identifier: u64,
    title_identifier: u64,
    caption_identifier: u64,
    movie_data_identifier: u64,
    geometry: Geometry,
    duration_seconds: f32,
    kind: MediaKind,
}

impl MediaArchiveWrite {
    /// Build a movie archive request.
    #[must_use]
    pub const fn movie(
        parent_identifier: u64,
        style_identifier: u64,
        title_identifier: u64,
        caption_identifier: u64,
        movie_data_identifier: u64,
        poster_data_identifier: Option<u64>,
        geometry: Geometry,
        duration_seconds: f32,
        natural_size: Size,
        poster_image_generated_with_alpha_support: bool,
    ) -> Self {
        Self {
            parent_identifier,
            style_identifier,
            title_identifier,
            caption_identifier,
            movie_data_identifier,
            geometry,
            duration_seconds,
            kind: MediaKind::movie(
                poster_data_identifier,
                natural_size,
                poster_image_generated_with_alpha_support,
            ),
        }
    }

    /// Build an audio-only archive request.
    #[must_use]
    pub const fn audio(
        parent_identifier: u64,
        style_identifier: u64,
        title_identifier: u64,
        caption_identifier: u64,
        movie_data_identifier: u64,
        geometry: Geometry,
        duration_seconds: f32,
    ) -> Self {
        Self {
            parent_identifier,
            style_identifier,
            title_identifier,
            caption_identifier,
            movie_data_identifier,
            geometry,
            duration_seconds,
            kind: MediaKind::Audio,
        }
    }

    /// Return the parent object identifier.
    #[must_use]
    pub const fn parent_identifier(self) -> u64 {
        self.parent_identifier
    }

    /// Return the style object identifier.
    #[must_use]
    pub const fn style_identifier(self) -> u64 {
        self.style_identifier
    }

    /// Return the title object identifier.
    #[must_use]
    pub const fn title_identifier(self) -> u64 {
        self.title_identifier
    }

    /// Return the caption object identifier.
    #[must_use]
    pub const fn caption_identifier(self) -> u64 {
        self.caption_identifier
    }

    /// Return the movie/audio data object identifier.
    #[must_use]
    pub const fn movie_data_identifier(self) -> u64 {
        self.movie_data_identifier
    }

    /// Return authored geometry.
    #[must_use]
    pub const fn geometry(self) -> Geometry {
        self.geometry
    }

    /// Return native duration in seconds.
    #[must_use]
    pub const fn duration_seconds(self) -> f32 {
        self.duration_seconds
    }

    /// Return the selected media kind.
    #[must_use]
    pub const fn kind(self) -> MediaKind {
        self.kind
    }
}

/// Finite resource policy for one fresh media archive encoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeOptions {
    max_output_bytes: usize,
    max_references: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_allocations: usize,
}

impl EncodeOptions {
    /// Construct an explicit output/reference/field/work/allocation policy.
    #[must_use]
    pub const fn new(
        max_output_bytes: usize,
        max_references: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_allocations: usize,
    ) -> Self {
        Self {
            max_output_bytes,
            max_references,
            max_fields,
            max_work_bytes,
            max_allocations,
        }
    }

    /// Build a conservative finite policy for a typed media request.
    #[must_use]
    pub const fn for_write(_write: &MediaArchiveWrite) -> Self {
        Self::new(
            MAX_DEFAULT_OUTPUT_BYTES,
            MAX_DEFAULT_REFERENCES,
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

    /// Replace the object/data-reference ceiling.
    #[must_use]
    pub const fn with_max_references(mut self, maximum: usize) -> Self {
        self.max_references = maximum;
        self
    }

    /// Replace the authored-field ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the aggregate measurement-and-write work ceiling.
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

/// Exact finite resource consumption of one encoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeReport {
    output_bytes: usize,
    references: usize,
    fields: usize,
    work_bytes: usize,
    allocations: usize,
}

impl EncodeReport {
    /// Return the exact encoded output length.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Return the number of object/data references emitted.
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    /// Return the number of emitted protobuf fields, including nested fields.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Return strict measurement plus encode work charged by this module.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Return the number of top-level output allocations charged.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

/// Encoded media payload plus exact resource evidence.
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

    /// Consume the output and return its encoded payload.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

/// Invalid value at the typed authoring boundary.
#[derive(Debug, Clone, Copy, PartialEq)]
#[non_exhaustive]
pub enum InvalidInput {
    /// A required object/data identifier was zero.
    ZeroIdentifier(&'static str),
    /// A floating-point value was NaN or infinite.
    NonFinite(&'static str),
    /// A size was negative.
    NegativeSize(&'static str),
    /// A duration was negative.
    NegativeDuration,
}

/// Finite resource exceeded by one media encoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncodeLimit {
    /// Encoded bytes exceed the caller's output ceiling.
    OutputBytes { observed: usize, maximum: usize },
    /// Object/data references exceed the caller's reference ceiling.
    References { observed: usize, maximum: usize },
    /// Authored fields exceed the caller's field ceiling.
    Fields { observed: usize, maximum: usize },
    /// Measurement plus write work exceeds the caller's work ceiling.
    WorkBytes { observed: usize, maximum: usize },
    /// Top-level allocation count exceeds the caller's allocation ceiling.
    Allocations { observed: usize, maximum: usize },
}

/// Failure from a typed, bounded media encoding.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum EncodeError {
    /// Input failed the strict scalar/reference policy.
    InvalidInput(InvalidInput),
    /// A finite caller budget was exceeded.
    Resource(EncodeLimit),
    /// The output vector could not be pre-reserved exactly.
    Allocation { amount: usize },
    /// Buffa rejected the generated projection or its bound.
    Buffa(buffa::EncodeError),
    /// The encoded length differed from the preflight measurement.
    Verification,
}

impl fmt::Display for EncodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidInput(input) => {
                write!(formatter, "invalid Keynote media input: {input:?}")
            },
            Self::Resource(_) => formatter.write_str("Keynote media encoding limit exceeded"),
            Self::Allocation { amount } => {
                write!(
                    formatter,
                    "Keynote media output allocation failed ({amount} bytes)"
                )
            },
            Self::Buffa(error) => error.fmt(formatter),
            Self::Verification => formatter.write_str("Keynote media encoding verification failed"),
        }
    }
}

impl std::error::Error for EncodeError {}

impl From<buffa::EncodeError> for EncodeError {
    fn from(error: buffa::EncodeError) -> Self {
        Self::Buffa(error)
    }
}

/// Encode one fresh Keynote audio/movie archive with an exact finite policy.
pub fn encode_media_archive(
    write: &MediaArchiveWrite,
    options: EncodeOptions,
) -> Result<Vec<u8>, EncodeError> {
    encode_media_archive_with_report(write, options).map(EncodeOutput::into_bytes)
}

/// Encode one fresh Keynote audio/movie archive and return exact resource
/// evidence.  Validation and preflight happen before any output allocation.
pub fn encode_media_archive_with_report(
    write: &MediaArchiveWrite,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    validate_write(write)?;

    let (references, fields, allocations) = media_shape(write);
    let output_bytes = media_encoded_len(write).ok_or(EncodeError::Verification)?;
    let work_bytes = output_bytes
        .checked_mul(3)
        .and_then(|work| work.checked_add(fields))
        .ok_or(EncodeError::Verification)?;
    let report = EncodeReport {
        output_bytes,
        references,
        fields,
        work_bytes,
        allocations,
    };
    preflight(report, options)?;

    // The static size is calculated before any `MessageFieldView::set` boxes
    // are created.  The generated view measurement below is an independent
    // parity check, not the resource preflight itself.
    let view = media_view(write);
    let measured_output_bytes =
        usize::try_from(view.try_encoded_len()?).map_err(|_| EncodeError::Verification)?;
    if measured_output_bytes != output_bytes {
        return Err(EncodeError::Verification);
    }

    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(output_bytes)
        .map_err(|_| EncodeError::Allocation {
            amount: output_bytes,
        })?;
    if bytes.capacity() < output_bytes {
        return Err(EncodeError::Allocation {
            amount: output_bytes,
        });
    }
    let maximum = u32::try_from(options.max_output_bytes).unwrap_or(u32::MAX);
    let encoded = view.try_encode_bounded(maximum, &mut bytes)?;
    if usize::try_from(encoded).ok() != Some(output_bytes) || bytes.len() != output_bytes {
        return Err(EncodeError::Verification);
    }
    Ok(EncodeOutput { bytes, report })
}

/// Canonical payload of a freshly allocated `TSD.StandinCaptionArchive`.
#[must_use]
pub const fn canonical_standin_payload() -> &'static [u8] {
    &[]
}

fn media_shape(write: &MediaArchiveWrite) -> (usize, usize, usize) {
    let (references, top_fields, poster) = match write.kind {
        MediaKind::Audio => (AUDIO_REFERENCES, 15, false),
        MediaKind::Movie {
            poster_data_identifier: Some(_),
            ..
        } => (MOVIE_REFERENCES_WITH_POSTER, 16, true),
        MediaKind::Movie {
            poster_data_identifier: None,
            ..
        } => (MOVIE_REFERENCES_WITHOUT_POSTER, 15, false),
    };
    let geometry_fields = 6
        + usize::from(write.geometry.flags.is_some())
        + usize::from(write.geometry.angle.is_some());
    // Top-level fields, drawable direct fields, geometry descendants,
    // exterior-wrap fields, reference identifier fields, data identifier
    // fields, and original/natural-size scalar fields.
    let fields = top_fields + 9 + geometry_fields + 6 + 4 + 1 + usize::from(poster) + 4;
    // MessageFieldView::set allocates one box for each populated nested view;
    // the final Vec allocation is charged separately.
    let nested_boxes = 12 + usize::from(poster);
    (references, fields, nested_boxes + 1)
}

fn media_encoded_len(write: &MediaArchiveWrite) -> Option<usize> {
    let geometry = len_field(1, {
        let mut length = len_field(1, 10)?; // Point
        length = length.checked_add(len_field(2, 10)?)?; // Size
        if let Some(flags) = write.geometry.flags {
            length = length.checked_add(varint_field(3, u64::from(flags)))?;
        }
        if write.geometry.angle.is_some() {
            length = length.checked_add(fixed32_field(4))?;
        }
        length
    })?;
    let drawable = geometry
        .checked_add(len_field(2, reference_len(write.parent_identifier))?)?
        .checked_add(len_field(3, exterior_text_wrap_len())?)?
        .checked_add(bool_field(5))?
        .checked_add(bool_field(7))?
        .checked_add(len_field(10, reference_len(write.title_identifier))?)?
        .checked_add(len_field(11, reference_len(write.caption_identifier))?)?
        .checked_add(bool_field(12))?
        .checked_add(bool_field(13))?;
    let mut length = len_field(1, drawable)?
        .checked_add(len_field(
            14,
            data_reference_len(write.movie_data_identifier),
        )?)?
        .checked_add(fixed32_field(3))?
        .checked_add(fixed32_field(4))?
        .checked_add(fixed32_field(5))?
        .checked_add(varint_field(24, 0))?
        .checked_add(fixed32_field(7))?
        .checked_add(bool_field(9))?
        .checked_add(bool_field(18))?
        .checked_add(bool_field(28))?;
    if let MediaKind::Movie {
        poster_data_identifier: Some(identifier),
        ..
    } = write.kind
    {
        length = length.checked_add(len_field(15, data_reference_len(identifier))?)?;
    }
    length = length
        .checked_add(bool_field(23))?
        .checked_add(varint_field(13, 0))?
        .checked_add(len_field(19, reference_len(write.style_identifier))?)?
        .checked_add(len_field(20, size_len(write_natural_size(write.kind)))?)?
        .checked_add(len_field(21, size_len(write_natural_size(write.kind)))?)?;
    Some(length)
}

fn write_natural_size(kind: MediaKind) -> Size {
    match kind {
        MediaKind::Audio => Size::new(0.0, 0.0),
        MediaKind::Movie { natural_size, .. } => natural_size,
    }
}

fn reference_len(identifier: u64) -> usize {
    varint_field(1, identifier)
}

fn data_reference_len(identifier: u64) -> usize {
    reference_len(identifier)
}

fn size_len(_size: Size) -> usize {
    fixed32_field(1) + fixed32_field(2)
}

fn exterior_text_wrap_len() -> usize {
    varint_field(1, 4)
        + varint_field(2, 2)
        + varint_field(3, 1)
        + fixed32_field(4)
        + fixed32_field(5)
        + bool_field(6)
}

fn key_len(field: u32, wire_type: u8) -> usize {
    varint_size((u64::from(field) << 3) | u64::from(wire_type))
}

fn varint_size(mut value: u64) -> usize {
    let mut size = 1;
    while value >= 0x80 {
        value >>= 7;
        size += 1;
    }
    size
}

fn varint_field(field: u32, value: u64) -> usize {
    key_len(field, 0) + varint_size(value)
}

fn bool_field(field: u32) -> usize {
    varint_field(field, 1)
}

fn fixed32_field(field: u32) -> usize {
    key_len(field, 5) + 4
}

fn len_field(field: u32, inner_length: usize) -> Option<usize> {
    key_len(field, 2)
        .checked_add(varint_size(u64::try_from(inner_length).ok()?))?
        .checked_add(inner_length)
}

fn validate_write(write: &MediaArchiveWrite) -> Result<(), EncodeError> {
    for (identifier, name) in [
        (write.parent_identifier, "parent_identifier"),
        (write.style_identifier, "style_identifier"),
        (write.title_identifier, "title_identifier"),
        (write.caption_identifier, "caption_identifier"),
        (write.movie_data_identifier, "movie_data_identifier"),
    ] {
        if identifier == 0 {
            return Err(EncodeError::InvalidInput(InvalidInput::ZeroIdentifier(
                name,
            )));
        }
    }

    let geometry = write.geometry;
    let point = geometry.position;
    for (value, name) in [
        (point.x, "geometry.position.x"),
        (point.y, "geometry.position.y"),
    ] {
        if !value.is_finite() {
            return Err(EncodeError::InvalidInput(InvalidInput::NonFinite(name)));
        }
    }
    validate_size(geometry.size, "geometry.size")?;
    if let Some(angle) = geometry.angle {
        if !angle.is_finite() {
            return Err(EncodeError::InvalidInput(InvalidInput::NonFinite(
                "geometry.angle",
            )));
        }
    }
    if !write.duration_seconds.is_finite() {
        return Err(EncodeError::InvalidInput(InvalidInput::NonFinite(
            "duration_seconds",
        )));
    }
    if write.duration_seconds < 0.0 {
        return Err(EncodeError::InvalidInput(InvalidInput::NegativeDuration));
    }

    if let MediaKind::Movie {
        poster_data_identifier,
        natural_size,
        ..
    } = write.kind
    {
        if let Some(identifier) = poster_data_identifier {
            if identifier == 0 {
                return Err(EncodeError::InvalidInput(InvalidInput::ZeroIdentifier(
                    "poster_data_identifier",
                )));
            }
        }
        validate_size(natural_size, "natural_size")?;
    }
    Ok(())
}

fn validate_size(size: Size, name: &'static str) -> Result<(), EncodeError> {
    for (value, field) in [
        (
            size.width,
            if name == "natural_size" {
                "natural_size.width"
            } else {
                "geometry.size.width"
            },
        ),
        (
            size.height,
            if name == "natural_size" {
                "natural_size.height"
            } else {
                "geometry.size.height"
            },
        ),
    ] {
        if !value.is_finite() {
            return Err(EncodeError::InvalidInput(InvalidInput::NonFinite(field)));
        }
        if value < 0.0 {
            return Err(EncodeError::InvalidInput(InvalidInput::NegativeSize(name)));
        }
    }
    Ok(())
}

fn preflight(report: EncodeReport, options: EncodeOptions) -> Result<(), EncodeError> {
    let limits = [
        (report.output_bytes > options.max_output_bytes).then_some(EncodeLimit::OutputBytes {
            observed: report.output_bytes,
            maximum: options.max_output_bytes,
        }),
        (report.references > options.max_references).then_some(EncodeLimit::References {
            observed: report.references,
            maximum: options.max_references,
        }),
        (report.fields > options.max_fields).then_some(EncodeLimit::Fields {
            observed: report.fields,
            maximum: options.max_fields,
        }),
        (report.work_bytes > options.max_work_bytes).then_some(EncodeLimit::WorkBytes {
            observed: report.work_bytes,
            maximum: options.max_work_bytes,
        }),
        (report.allocations > options.max_allocations).then_some(EncodeLimit::Allocations {
            observed: report.allocations,
            maximum: options.max_allocations,
        }),
    ];
    if let Some(limit) = limits.into_iter().flatten().next() {
        return Err(EncodeError::Resource(limit));
    }
    Ok(())
}

fn reference(identifier: u64) -> buffa_tsd::ReferenceView<'static> {
    buffa_tsd::ReferenceView {
        identifier,
        ..Default::default()
    }
}

fn data_reference(identifier: u64) -> buffa_tsd::DataReferenceView<'static> {
    buffa_tsd::DataReferenceView {
        identifier,
        ..Default::default()
    }
}

fn point(value: Point) -> buffa_tsd::PointView<'static> {
    buffa_tsd::PointView {
        x: value.x,
        y: value.y,
        ..Default::default()
    }
}

fn size(value: Size) -> buffa_tsd::SizeView<'static> {
    buffa_tsd::SizeView {
        width: value.width,
        height: value.height,
        ..Default::default()
    }
}

fn media_view(write: &MediaArchiveWrite) -> buffa_tsd::MovieArchiveView<'static> {
    let geometry = buffa_tsd::GeometryArchiveView {
        position: buffa::MessageFieldView::set(point(write.geometry.position)),
        size: buffa::MessageFieldView::set(size(write.geometry.size)),
        flags: write.geometry.flags,
        angle: write.geometry.angle,
    };
    let exterior_text_wrap = buffa_tsd::ExteriorTextWrapArchiveView {
        r#type: Some(4),
        direction: Some(2),
        fit_type: Some(1),
        margin: Some(12.0),
        alpha_threshold: Some(0.5),
        is_html_wrap: Some(false),
        ..Default::default()
    };
    let drawable = buffa_tsd::DrawableArchiveView {
        geometry: buffa::MessageFieldView::set(geometry),
        parent: buffa::MessageFieldView::set(reference(write.parent_identifier)),
        exterior_text_wrap: buffa::MessageFieldView::set(exterior_text_wrap),
        locked: Some(false),
        aspect_ratio_locked: Some(true),
        title: buffa::MessageFieldView::set(reference(write.title_identifier)),
        caption: buffa::MessageFieldView::set(reference(write.caption_identifier)),
        title_hidden: Some(false),
        caption_hidden: Some(false),
    };
    let (audio_only, natural_size, poster_data_identifier, alpha_support) = match write.kind {
        MediaKind::Audio => (true, Size::new(0.0, 0.0), None, false),
        MediaKind::Movie {
            poster_data_identifier,
            natural_size,
            poster_image_generated_with_alpha_support,
        } => (
            false,
            natural_size,
            poster_data_identifier,
            poster_image_generated_with_alpha_support,
        ),
    };
    let poster_image_data = match poster_data_identifier {
        Some(identifier) => buffa::MessageFieldView::set(data_reference(identifier)),
        None => buffa::MessageFieldView::unset(),
    };
    buffa_tsd::MovieArchiveView {
        super_: buffa::MessageFieldView::set(drawable),
        movie_data: buffa::MessageFieldView::set(data_reference(write.movie_data_identifier)),
        start_time: Some(0.0),
        end_time: Some(write.duration_seconds),
        poster_time: Some(0.0),
        loop_option: Some(0),
        volume: Some(1.0),
        audio_only: Some(audio_only),
        streaming: Some(false),
        plays_across_slides: Some(true),
        poster_image_data,
        poster_image_generated_with_alpha_support: Some(alpha_support),
        flags: Some(0),
        style: buffa::MessageFieldView::set(reference(write.style_identifier)),
        original_size: buffa::MessageFieldView::set(size(natural_size)),
        natural_size: buffa::MessageFieldView::set(size(natural_size)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use prost::Message as _;

    fn geometry() -> Geometry {
        Geometry::new(
            Point::new(12.5, -4.25),
            Size::new(320.0, 180.0),
            Some(0),
            Some(0.0),
        )
    }

    fn movie_write(poster_data_identifier: Option<u64>) -> MediaArchiveWrite {
        MediaArchiveWrite::movie(
            11,
            17,
            19,
            23,
            29,
            poster_data_identifier,
            geometry(),
            7.5,
            Size::new(1920.0, 1080.0),
            true,
        )
    }

    fn audio_write() -> MediaArchiveWrite {
        MediaArchiveWrite::audio(11, 17, 19, 23, 29, geometry(), 7.5)
    }

    fn native_oracle(write: &MediaArchiveWrite) -> Vec<u8> {
        let drawable = crate::tsd::DrawableArchive {
            geometry: Some(crate::tsd::GeometryArchive {
                position: Some(crate::tsp::Point {
                    x: write.geometry.position.x,
                    y: write.geometry.position.y,
                }),
                size: Some(crate::tsp::Size {
                    width: write.geometry.size.width,
                    height: write.geometry.size.height,
                }),
                flags: write.geometry.flags,
                angle: write.geometry.angle,
            }),
            parent: Some(crate::tsp::Reference {
                identifier: write.parent_identifier,
                ..Default::default()
            }),
            exterior_text_wrap: Some(crate::tsd::ExteriorTextWrapArchive {
                r#type: Some(4),
                direction: Some(2),
                fit_type: Some(1),
                margin: Some(12.0),
                alpha_threshold: Some(0.5),
                is_html_wrap: Some(false),
            }),
            locked: Some(false),
            aspect_ratio_locked: Some(true),
            title: Some(crate::tsp::Reference {
                identifier: write.title_identifier,
                ..Default::default()
            }),
            caption: Some(crate::tsp::Reference {
                identifier: write.caption_identifier,
                ..Default::default()
            }),
            title_hidden: Some(false),
            caption_hidden: Some(false),
            ..Default::default()
        };
        let (audio_only, natural_size, poster_data_identifier, alpha_support) = match write.kind {
            MediaKind::Audio => (true, Size::new(0.0, 0.0), None, false),
            MediaKind::Movie {
                poster_data_identifier,
                natural_size,
                poster_image_generated_with_alpha_support,
            } => (
                false,
                natural_size,
                poster_data_identifier,
                poster_image_generated_with_alpha_support,
            ),
        };
        crate::tsd::MovieArchive {
            super_: drawable,
            movie_data: Some(crate::tsp::DataReference {
                identifier: write.movie_data_identifier,
            }),
            start_time: Some(0.0),
            end_time: Some(write.duration_seconds),
            poster_time: Some(0.0),
            loop_option: Some(crate::tsd::movie_archive::MovieLoopOption::None as i32),
            volume: Some(1.0),
            audio_only: Some(audio_only),
            streaming: Some(false),
            plays_across_slides: Some(true),
            poster_image_data: poster_data_identifier
                .map(|identifier| crate::tsp::DataReference { identifier }),
            poster_image_generated_with_alpha_support: Some(alpha_support),
            flags: Some(0),
            style: Some(crate::tsp::Reference {
                identifier: write.style_identifier,
                ..Default::default()
            }),
            original_size: Some(crate::tsp::Size {
                width: natural_size.width,
                height: natural_size.height,
            }),
            natural_size: Some(crate::tsp::Size {
                width: natural_size.width,
                height: natural_size.height,
            }),
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn read_varint(source: &[u8], cursor: &mut usize) -> Option<u64> {
        let mut value = 0u64;
        for shift in (0..70).step_by(7) {
            let byte = *source.get(*cursor)?;
            *cursor += 1;
            value |= u64::from(byte & 0x7f).checked_shl(u32::try_from(shift).ok()?)?;
            if byte & 0x80 == 0 {
                return Some(value);
            }
        }
        None
    }

    // Every length-delimited field in this private projection is another
    // projection message (there are no bytes/string fields).  Counting the
    // emitted wire fields recursively gives an independent census for the
    // report constants and catches a future field omission in preflight.
    fn wire_field_count(source: &[u8]) -> Option<usize> {
        let mut cursor = 0usize;
        let mut count = 0usize;
        while cursor < source.len() {
            let tag = read_varint(source, &mut cursor)?;
            let wire_type = u8::try_from(tag & 0x07).ok()?;
            count = count.checked_add(1)?;
            match wire_type {
                0 => {
                    let _ = read_varint(source, &mut cursor)?;
                },
                1 => cursor = cursor.checked_add(8)?,
                2 => {
                    let length = usize::try_from(read_varint(source, &mut cursor)?).ok()?;
                    let end = cursor.checked_add(length)?;
                    count = count.checked_add(wire_field_count(source.get(cursor..end)?)?)?;
                    cursor = end;
                },
                5 => cursor = cursor.checked_add(4)?,
                _ => return None,
            }
            if cursor > source.len() {
                return None;
            }
        }
        Some(count)
    }

    #[test]
    fn movie_projection_matches_native_with_poster() {
        let write = movie_write(Some(31));
        let encoded = encode_media_archive(&write, EncodeOptions::for_write(&write)).unwrap();
        assert_eq!(encoded, native_oracle(&write));
    }

    #[test]
    fn movie_projection_matches_native_without_poster() {
        let write = movie_write(None);
        let encoded = encode_media_archive(&write, EncodeOptions::for_write(&write)).unwrap();
        assert_eq!(encoded, native_oracle(&write));
    }

    #[test]
    fn audio_projection_matches_native() {
        let write = audio_write();
        let encoded = encode_media_archive(&write, EncodeOptions::for_write(&write)).unwrap();
        assert_eq!(encoded, native_oracle(&write));
    }

    #[test]
    fn optional_geometry_scalars_preserve_absent_wire_fields() {
        let write = MediaArchiveWrite::audio(
            11,
            17,
            19,
            23,
            29,
            Geometry::new(Point::new(12.5, -4.25), Size::new(320.0, 180.0), None, None),
            7.5,
        );
        let encoded = encode_media_archive(&write, EncodeOptions::for_write(&write)).unwrap();
        assert_eq!(encoded, native_oracle(&write));
    }

    #[test]
    fn invalid_zero_identifier_is_rejected_before_encoding() {
        let write = MediaArchiveWrite::audio(0, 17, 19, 23, 29, geometry(), 7.5);
        assert!(matches!(
            encode_media_archive(&write, EncodeOptions::for_write(&write)),
            Err(EncodeError::InvalidInput(InvalidInput::ZeroIdentifier(
                "parent_identifier"
            )))
        ));
    }

    #[test]
    fn non_finite_and_negative_geometry_are_rejected() {
        let non_finite = MediaArchiveWrite::audio(
            11,
            17,
            19,
            23,
            29,
            Geometry::new(
                Point::new(f32::NAN, 0.0),
                Size::new(1.0, 1.0),
                Some(0),
                Some(0.0),
            ),
            1.0,
        );
        assert!(matches!(
            encode_media_archive(&non_finite, EncodeOptions::for_write(&non_finite)),
            Err(EncodeError::InvalidInput(InvalidInput::NonFinite(
                "geometry.position.x"
            )))
        ));

        let negative = MediaArchiveWrite::audio(
            11,
            17,
            19,
            23,
            29,
            Geometry::new(
                Point::new(0.0, 0.0),
                Size::new(-1.0, 1.0),
                Some(0),
                Some(0.0),
            ),
            1.0,
        );
        assert!(matches!(
            encode_media_archive(&negative, EncodeOptions::for_write(&negative)),
            Err(EncodeError::InvalidInput(InvalidInput::NegativeSize(
                "geometry.size"
            )))
        ));
    }

    #[test]
    fn maximum_identifier_and_zero_sized_audio_are_supported() {
        let write = MediaArchiveWrite::audio(
            u64::MAX,
            u64::MAX - 1,
            u64::MAX - 2,
            u64::MAX - 3,
            u64::MAX - 4,
            Geometry::new(
                Point::new(0.0, 0.0),
                Size::new(0.0, 0.0),
                Some(0),
                Some(0.0),
            ),
            0.0,
        );
        let encoded = encode_media_archive(&write, EncodeOptions::for_write(&write)).unwrap();
        assert_eq!(encoded, native_oracle(&write));
    }

    #[test]
    fn exact_resource_limits_are_admitted_and_one_below_is_rejected() {
        let write = movie_write(Some(31));
        let output =
            encode_media_archive_with_report(&write, EncodeOptions::for_write(&write)).unwrap();
        let report = output.report();
        assert!(
            encode_media_archive(
                &write,
                EncodeOptions::for_write(&write).with_max_output_bytes(report.output_bytes())
            )
            .is_ok()
        );
        assert!(matches!(
            encode_media_archive(
                &write,
                EncodeOptions::for_write(&write).with_max_output_bytes(report.output_bytes() - 1)
            ),
            Err(EncodeError::Resource(EncodeLimit::OutputBytes { .. }))
        ));
        assert!(
            encode_media_archive(
                &write,
                EncodeOptions::for_write(&write).with_max_references(report.references())
            )
            .is_ok()
        );
        assert!(matches!(
            encode_media_archive(
                &write,
                EncodeOptions::for_write(&write).with_max_references(report.references() - 1)
            ),
            Err(EncodeError::Resource(EncodeLimit::References { .. }))
        ));
        assert!(
            encode_media_archive(
                &write,
                EncodeOptions::for_write(&write).with_max_fields(report.fields())
            )
            .is_ok()
        );
        assert!(matches!(
            encode_media_archive(
                &write,
                EncodeOptions::for_write(&write).with_max_fields(report.fields() - 1)
            ),
            Err(EncodeError::Resource(EncodeLimit::Fields { .. }))
        ));
        assert!(
            encode_media_archive(
                &write,
                EncodeOptions::for_write(&write).with_max_work_bytes(report.work_bytes())
            )
            .is_ok()
        );
        assert!(matches!(
            encode_media_archive(
                &write,
                EncodeOptions::for_write(&write).with_max_work_bytes(report.work_bytes() - 1)
            ),
            Err(EncodeError::Resource(EncodeLimit::WorkBytes { .. }))
        ));
        assert!(
            encode_media_archive(
                &write,
                EncodeOptions::for_write(&write).with_max_allocations(report.allocations())
            )
            .is_ok()
        );
        assert!(matches!(
            encode_media_archive(
                &write,
                EncodeOptions::for_write(&write).with_max_allocations(report.allocations() - 1)
            ),
            Err(EncodeError::Resource(EncodeLimit::Allocations { .. }))
        ));
    }

    #[test]
    fn report_field_census_matches_emitted_wire() {
        for write in [movie_write(Some(31)), movie_write(None), audio_write()] {
            let output =
                encode_media_archive_with_report(&write, EncodeOptions::for_write(&write)).unwrap();
            assert_eq!(
                output.report().fields(),
                wire_field_count(output.bytes()).expect("valid projection wire")
            );
        }
    }
}
