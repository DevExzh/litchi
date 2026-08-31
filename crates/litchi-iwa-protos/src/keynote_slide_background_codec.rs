//! Strict borrowed Keynote slide-background fill codec.
//!
//! A slide style stores its background as a complete `TSD.FillArchive`.
//! This module owns the wire boundary for that payload: known proto2 fields
//! are checked for canonical framing, singularity, finite values, required
//! nested values, and bounded resource use before a private Buffa lazy view is
//! forced.  The source slice remains authoritative for every rewrite.  In
//! particular, gradient/image branches and unknown extension fields are never
//! converted into generated owned messages.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict wire helpers remain beside the raw-preserving rewrite model."
)]

use core::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_slide_background_generated::LitchiIwaKeynoteBackgroundProjection as projection;

const FILL_COLOR_FIELD: u32 = 1;
const FILL_GRADIENT_FIELD: u32 = 2;
const FILL_IMAGE_FIELD: u32 = 3;
const COLOR_MODEL_FIELD: u32 = 1;
const COLOR_RED_FIELD: u32 = 3;
const COLOR_GREEN_FIELD: u32 = 4;
const COLOR_BLUE_FIELD: u32 = 5;
const COLOR_ALPHA_FIELD: u32 = 6;
const COLOR_CYAN_FIELD: u32 = 7;
const COLOR_MAGENTA_FIELD: u32 = 8;
const COLOR_YELLOW_FIELD: u32 = 9;
const COLOR_BLACK_FIELD: u32 = 10;
const COLOR_WHITE_FIELD: u32 = 11;
const COLOR_SPACE_FIELD: u32 = 12;
const GRADIENT_TYPE_FIELD: u32 = 1;
const GRADIENT_STOP_FIELD: u32 = 2;
const GRADIENT_OPACITY_FIELD: u32 = 3;
const GRADIENT_ADVANCED_FIELD: u32 = 4;
const GRADIENT_ANGLE_FIELD: u32 = 5;
const GRADIENT_TRANSFORM_FIELD: u32 = 6;
const STOP_COLOR_FIELD: u32 = 1;
const STOP_FRACTION_FIELD: u32 = 2;
const STOP_INFLECTION_FIELD: u32 = 3;
const ANGLE_RADIANS_FIELD: u32 = 2;
const TRANSFORM_START_FIELD: u32 = 1;
const TRANSFORM_END_FIELD: u32 = 2;
const TRANSFORM_SIZE_FIELD: u32 = 3;
const IMAGE_DATABASE_DATA_FIELD: u32 = 1;
const IMAGE_TECHNIQUE_FIELD: u32 = 2;
const IMAGE_TINT_FIELD: u32 = 3;
const IMAGE_SIZE_FIELD: u32 = 4;
const IMAGE_DATABASE_ORIGINAL_FIELD: u32 = 5;
const IMAGE_DATA_FIELD: u32 = 6;
const IMAGE_ORIGINAL_DATA_FIELD: u32 = 7;
const IMAGE_UNTAGGED_FIELD: u32 = 8;
const IMAGE_REFERENCE_COLOR_FIELD: u32 = 9;
const RESOURCE_IDENTIFIER_FIELD: u32 = 1;
const POINT_X_FIELD: u32 = 1;
const POINT_Y_FIELD: u32 = 2;
const SIZE_WIDTH_FIELD: u32 = 1;
const SIZE_HEIGHT_FIELD: u32 = 2;
const RGB_MODEL: i32 = 1;
const SRGB_SPACE: i32 = 1;
const P3_SPACE: i32 = 2;
const MAX_RECURSION_LIMIT: u32 = 64;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Finite resource profile for one complete native fill payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build an explicit finite profile.
    #[must_use]
    pub const fn new(max_message_bytes: usize, recursion_limit: u32) -> Self {
        Self {
            max_message_bytes,
            max_fields: max_message_bytes,
            max_work_bytes: max_message_bytes.saturating_mul(8),
            recursion_limit,
        }
    }

    /// Build a conservative profile for a borrowed source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        // A valid gradient is scanned once per nested preflight, once by the
        // stop iterator, and once again by the projection cross-check.  The
        // handwritten budget charges each pass at two bytes per source byte;
        // eight source lengths is therefore too tight for a real gradient.
        // Keep this convenience profile finite but leave enough headroom for
        // the complete strict/projection traversal.  Callers with a shared
        // aggregate budget should continue to use `new` plus
        // `with_resource_limits`.
        Self::new(bytes, 16).with_resource_limits(bytes, bytes.saturating_mul(16))
    }

    /// Override exact aggregate field and work ceilings.
    #[must_use]
    pub const fn with_resource_limits(mut self, max_fields: usize, max_work_bytes: usize) -> Self {
        self.max_fields = max_fields;
        self.max_work_bytes = max_work_bytes;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_message_bytes)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// RGB color space understood by the native Keynote fill writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RgbSpace {
    Srgb,
    DisplayP3,
}

/// Primitive validated color used by solid and gradient writes.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ColorWrite {
    pub red: f32,
    pub green: f32,
    pub blue: f32,
    pub alpha: f32,
    pub rgb_space: RgbSpace,
}

impl ColorWrite {
    #[must_use]
    pub const fn new(red: f32, green: f32, blue: f32, alpha: f32, rgb_space: RgbSpace) -> Self {
        Self {
            red,
            green,
            blue,
            alpha,
            rgb_space,
        }
    }
}

/// Gradient kind written by the native Keynote host.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GradientKind {
    Linear,
    Radial,
}

/// One native gradient stop. Values use the exact fraction/inflection
/// scalars consumed by `TSD.GradientArchive`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct GradientStopWrite {
    pub color: ColorWrite,
    pub fraction: f32,
    pub inflection: f32,
}

/// Typed gradient write, encoded with the same fields as Keynote's native
/// gradient writer. The stop slice is borrowed for the duration of rewrite.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct GradientWrite<'source> {
    pub kind: GradientKind,
    pub stops: &'source [GradientStopWrite],
    pub opacity: f32,
    pub advanced: bool,
    pub angle_radians: f32,
}

/// Typed replacement accepted by the raw-preserving fill codec.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BackgroundWrite<'source> {
    /// Emit an explicit empty `TSD.FillArchive` (native “No Fill”).
    Clear,
    /// Replace with a native RGB solid color.
    Solid(ColorWrite),
    /// Replace with a native linear/radial gradient.
    Gradient(GradientWrite<'source>),
    /// Validate and install an already encoded fill payload verbatim.
    Raw(&'source [u8]),
}

/// Borrowed RGB values selected from a native color message.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ColorSnapshot {
    pub model: i32,
    pub red: Option<f32>,
    pub green: Option<f32>,
    pub blue: Option<f32>,
    pub alpha: Option<f32>,
    pub cyan: Option<f32>,
    pub magenta: Option<f32>,
    pub yellow: Option<f32>,
    pub black: Option<f32>,
    pub white: Option<f32>,
    pub rgb_space: Option<i32>,
}

/// Borrowed gradient scalars. Stops are re-read from `raw` by
/// [`GradientSnapshot::stops`], so
/// decoding this snapshot never allocates a repeated generated view.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct GradientSnapshot<'source> {
    pub raw: &'source [u8],
    pub gradient_type: Option<i32>,
    pub opacity: Option<f32>,
    pub advanced: Option<bool>,
    pub angle_radians: Option<f32>,
    pub transform: Option<&'source [u8]>,
}

/// One borrowed gradient stop yielded by [`GradientSnapshot::stops`].
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct GradientStopSnapshot {
    pub color: Option<ColorSnapshot>,
    pub fraction: Option<f32>,
    pub inflection: Option<f32>,
}

/// Borrowed iterator over repeated gradient stop payloads.
pub struct GradientStops<'source> {
    input: &'source [u8],
    options: DecodeOptions,
}

impl<'source> Iterator for GradientStops<'source> {
    type Item = Result<GradientStopSnapshot, DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            let field = match next_field(&mut self.input) {
                Ok(Some(field)) => field,
                Ok(None) => return None,
                Err(error) => return Some(Err(error)),
            };
            if field.number != GRADIENT_STOP_FIELD {
                continue;
            }
            let payload = match field.bytes() {
                Ok(payload) => payload,
                Err(error) => return Some(Err(error)),
            };
            return Some(parse_gradient_stop(payload, self.options));
        }
    }
}

impl<'source> GradientSnapshot<'source> {
    /// Iterate validated native gradient stops without retaining a vector.
    #[must_use]
    pub fn stops(self, options: DecodeOptions) -> GradientStops<'source> {
        GradientStops {
            input: self.raw,
            options,
        }
    }
}

/// Borrowed image-fill scalars and validated resource identifiers.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ImageSnapshot<'source> {
    pub raw: &'source [u8],
    pub database_image_identifier: Option<u64>,
    pub technique: Option<i32>,
    pub tint: Option<ColorSnapshot>,
    pub fill_size: Option<SizeSnapshot>,
    pub database_original_identifier: Option<u64>,
    pub image_identifier: Option<u64>,
    pub original_image_identifier: Option<u64>,
    pub interprets_untagged_image_data_as_generic: Option<bool>,
    pub reference_color: Option<ColorSnapshot>,
}

/// Required finite native point/size values selected by a fill transform.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SizeSnapshot {
    pub width: f32,
    pub height: f32,
}

/// Borrowed classification of one complete native `TSD.FillArchive`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BackgroundSnapshot<'source> {
    None {
        raw: &'source [u8],
    },
    Solid {
        raw: &'source [u8],
        color: ColorSnapshot,
    },
    Gradient {
        raw: &'source [u8],
        gradient: GradientSnapshot<'source>,
    },
    Image {
        raw: &'source [u8],
        image: ImageSnapshot<'source>,
    },
    Opaque {
        raw: &'source [u8],
    },
}

impl<'source> BackgroundSnapshot<'source> {
    /// Return the exact caller-owned fill bytes accepted by the codec.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        match self {
            Self::None { raw }
            | Self::Solid { raw, .. }
            | Self::Gradient { raw, .. }
            | Self::Image { raw, .. }
            | Self::Opaque { raw } => raw,
        }
    }
}

/// Finite counters for one strict scan and its bounded projection checks.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    pub input_bytes: usize,
    pub fields: usize,
    pub work_bytes: usize,
    pub max_depth: u32,
}

/// Finite counters for one source-preserving rewrite and strict readback.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    pub input_bytes: usize,
    pub output_bytes: usize,
    pub fields: usize,
    pub work_bytes: usize,
    pub max_depth: u32,
    pub changed: bool,
    /// Handwritten output-buffer reservations; generated Buffa internals are
    /// intentionally outside this counter.
    pub allocations: usize,
}

struct OutputMeasure {
    bytes: usize,
    scan_bytes: usize,
}

/// Byte or nesting resource failure from strict preflight.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    Bytes { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    Resource(WireResourceLimit),
    Field { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Allocation { requested: usize },
    Projection,
}

/// Strict wire, semantic, projection, or resource failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

impl DecodeError {
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

    const fn resource(limit: WireResourceLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Resource(limit),
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    /// Return byte/nesting resource detail when applicable.
    #[must_use]
    pub const fn wire_resource_limit(&self) -> Option<WireResourceLimit> {
        match self.kind {
            DecodeErrorKind::Resource(limit) => Some(limit),
            _ => None,
        }
    }

    /// Return exact observed and allowed field counts when applicable.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Field { observed, maximum } => Some((observed, maximum)),
            _ => None,
        }
    }

    /// Return exact observed and allowed work bytes when applicable.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Work { observed, maximum } => Some((observed, maximum)),
            _ => None,
        }
    }

    /// Return the requested allocation when a fallible rewrite buffer could
    /// not be reserved.
    #[must_use]
    pub const fn allocation_requested(&self) -> Option<usize> {
        match self.kind {
            DecodeErrorKind::Allocation { requested } => Some(requested),
            _ => None,
        }
    }

    /// Return the selected duplicate field description.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            _ => None,
        }
    }

    /// Return the selected missing required field description.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            _ => None,
        }
    }

    /// Return a stable canonicality explanation.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            _ => None,
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
            DecodeErrorKind::Resource(WireResourceLimit::Bytes { .. }) => {
                formatter.write_str("Keynote slide-background byte limit exceeded")
            },
            DecodeErrorKind::Resource(WireResourceLimit::Nesting { .. }) => {
                formatter.write_str("Keynote slide-background nesting limit exceeded")
            },
            DecodeErrorKind::Field { observed, maximum } => {
                write!(formatter, "visited {observed} fields; maximum is {maximum}")
            },
            DecodeErrorKind::Work { observed, maximum } => write!(
                formatter,
                "requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Allocation { requested } => {
                write!(formatter, "could not reserve {requested} bytes")
            },
            DecodeErrorKind::Projection => formatter.write_str(
                "Keynote slide-background strict preflight disagrees with Buffa projection",
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

/// Decode and cross-check one complete native fill payload.
pub fn decode_slide_background<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<BackgroundSnapshot<'source>, DecodeError> {
    Ok(decode_slide_background_with_report(source, options)?.0)
}

/// Decode one fill and return exact finite resource consumption.
pub fn decode_slide_background_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(BackgroundSnapshot<'source>, DecodeReport), DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_fill(source, options, &mut budget, 0)?;
    let view: projection::KeynoteSlideBackgroundArchiveLazyView<'source> =
        options.buffa().decode_lazy_view(source)?;
    cross_check_projection(&view, &strict, options, &mut budget)?;
    Ok((strict, budget.report(source.len())))
}

/// Rewrite a fill with a typed clear/solid/gradient or validated raw payload.
/// The candidate is strictly decoded and cross-checked before publication.
pub fn rewrite_slide_background_with_report<'source>(
    source: &'source [u8],
    write: BackgroundWrite<'source>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let (snapshot, input_report) = decode_slide_background_with_report(source, options)?;
    let measured_output = measure_write_output_len(source, write)?;
    let estimated_output = measured_output.bytes;
    if estimated_output > options.max_message_bytes {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: estimated_output,
            maximum: options.max_message_bytes,
        }));
    }
    let measure_work = measured_output
        .scan_bytes
        .checked_mul(2)
        .ok_or_else(DecodeError::projection)?;
    let minimum_output_work = estimated_output
        .checked_mul(2)
        .ok_or_else(DecodeError::projection)?;
    let minimum_total_work = input_report
        .work_bytes
        .checked_add(measure_work)
        .and_then(|value| value.checked_add(minimum_output_work))
        .ok_or_else(DecodeError::projection)?;
    if minimum_total_work > options.max_work_bytes {
        return Err(DecodeError {
            kind: DecodeErrorKind::Work {
                observed: minimum_total_work,
                maximum: options.max_work_bytes,
            },
        });
    }
    let replacement = encode_write(source, snapshot, write, options.max_message_bytes)?;
    let changed = replacement.as_slice() != source;
    let candidate = replacement;
    if candidate.len() > options.max_message_bytes {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: candidate.len(),
            maximum: options.max_message_bytes,
        }));
    }
    let (readback, readback_report) = decode_slide_background_with_report(&candidate, options)?;
    ensure_write_matches(write, &readback, options)?;
    // A rewrite has one aggregate resource budget.  The candidate readback
    // uses the same per-message limits as an ordinary decode, then this final
    // check accounts for the source scan, encoded candidate bytes, and the
    // candidate scan together.  Encoding work is charged at one byte per
    // emitted byte; handwritten preflight scans charge two per source byte.
    let fields = input_report
        .fields
        .checked_add(readback_report.fields)
        .ok_or_else(DecodeError::projection)?;
    if fields > options.max_fields {
        return Err(DecodeError {
            kind: DecodeErrorKind::Field {
                observed: fields,
                maximum: options.max_fields,
            },
        });
    }
    let work_bytes = input_report
        .work_bytes
        .checked_add(measure_work)
        .and_then(|value| value.checked_add(candidate.len()))
        .and_then(|value| value.checked_add(readback_report.work_bytes))
        .ok_or_else(DecodeError::projection)?;
    let allocations = count_rewrite_allocations(source, write)?;
    if work_bytes > options.max_work_bytes {
        return Err(DecodeError {
            kind: DecodeErrorKind::Work {
                observed: work_bytes,
                maximum: options.max_work_bytes,
            },
        });
    }
    Ok((
        candidate,
        RewriteReport {
            input_bytes: input_report.input_bytes,
            output_bytes: readback_report.input_bytes,
            fields,
            work_bytes,
            max_depth: input_report.max_depth.max(readback_report.max_depth),
            changed,
            allocations,
        },
    ))
}

/// Rewrite without exposing the resource report.
pub fn rewrite_slide_background<'source>(
    source: &'source [u8],
    write: BackgroundWrite<'source>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    Ok(rewrite_slide_background_with_report(source, write, options)?.0)
}

struct Budget {
    fields: usize,
    work: usize,
    max_fields: usize,
    max_work: usize,
    max_depth: u32,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work: 0,
            max_fields: options.max_fields,
            max_work: options.max_work_bytes,
            max_depth: 0,
        }
    }

    fn charge(
        &mut self,
        source: &[u8],
        options: DecodeOptions,
        depth: u32,
    ) -> Result<(), DecodeError> {
        self.max_depth = self.max_depth.max(depth);
        if depth > options.recursion_limit {
            return Err(DecodeError::resource(WireResourceLimit::Nesting {
                observed: depth,
                maximum: options.recursion_limit,
            }));
        }
        let work = source
            .len()
            .checked_mul(2)
            .and_then(|cost| self.work.checked_add(cost))
            .ok_or_else(DecodeError::projection)?;
        if work > self.max_work {
            return Err(DecodeError {
                kind: DecodeErrorKind::Work {
                    observed: work,
                    maximum: self.max_work,
                },
            });
        }
        self.work = work;
        let mut input = source;
        while let Some(_field) = next_field(&mut input)? {
            let observed = self
                .fields
                .checked_add(1)
                .ok_or_else(DecodeError::projection)?;
            if observed > self.max_fields {
                return Err(DecodeError {
                    kind: DecodeErrorKind::Field {
                        observed,
                        maximum: self.max_fields,
                    },
                });
            }
            self.fields = observed;
        }
        Ok(())
    }

    fn report(&self, input_bytes: usize) -> DecodeReport {
        DecodeReport {
            input_bytes,
            fields: self.fields,
            work_bytes: self.work,
            max_depth: self.max_depth,
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum Value<'source> {
    Varint(u64, bool),
    Fixed32(u32),
    Fixed64,
    Bytes(&'source [u8], bool),
}

#[derive(Clone, Copy, Debug)]
struct Field<'source> {
    number: u32,
    wire: u8,
    value: Value<'source>,
    canonical_key: bool,
}

impl<'source> Field<'source> {
    fn varint(self) -> Result<u64, DecodeError> {
        self.require_wire(0)?;
        match self.value {
            Value::Varint(value, canonical) if canonical => Ok(value),
            Value::Varint(_, false) => Err(DecodeError::noncanonical("protobuf varint value")),
            _ => Err(DecodeError::projection()),
        }
    }

    fn int32(self) -> Result<i32, DecodeError> {
        let value = self.varint()?;
        if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
            return Err(DecodeError::noncanonical(
                "int32 scalar is not sign-extended",
            ));
        }
        #[allow(
            clippy::cast_possible_wrap,
            reason = "Canonical int32 bounds were checked above."
        )]
        Ok(value as i32)
    }

    fn bool(self) -> Result<bool, DecodeError> {
        match self.varint()? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
        }
    }

    fn bytes(self) -> Result<&'source [u8], DecodeError> {
        self.require_wire(2)?;
        match self.value {
            Value::Bytes(value, canonical) if canonical => Ok(value),
            Value::Bytes(_, false) => Err(DecodeError::noncanonical("length-delimited size")),
            _ => Err(DecodeError::projection()),
        }
    }

    fn fixed32(self) -> Result<u32, DecodeError> {
        self.require_wire(5)?;
        match self.value {
            Value::Fixed32(value) => Ok(value),
            _ => Err(DecodeError::projection()),
        }
    }

    fn float(self) -> Result<f32, DecodeError> {
        let value = f32::from_bits(self.fixed32()?);
        if !value.is_finite() {
            return Err(DecodeError::noncanonical("float scalar is not finite"));
        }
        Ok(value)
    }

    fn require_wire(self, expected: u8) -> Result<(), DecodeError> {
        if !self.canonical_key {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
        if self.wire != expected {
            return Err(DecodeError::from(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected,
                actual: self.wire,
            }));
        }
        Ok(())
    }
}

fn next_field<'source>(input: &mut &'source [u8]) -> Result<Option<Field<'source>>, DecodeError> {
    if input.is_empty() {
        return Ok(None);
    }
    let (tag, canonical_key) = take_varint(input)?;
    let tag = u32::try_from(tag).map_err(|_error| buffa::DecodeError::InvalidFieldNumber)?;
    let number = tag >> 3;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let wire = (tag & 7) as u8;
    let value = match wire {
        0 => {
            let (value, canonical) = take_varint(input)?;
            Value::Varint(value, canonical)
        },
        1 => {
            take_exact(input, 8)?;
            Value::Fixed64
        },
        2 => {
            let (length, canonical) = take_varint(input)?;
            let length =
                usize::try_from(length).map_err(|_error| buffa::DecodeError::MessageTooLarge)?;
            Value::Bytes(take_exact(input, length)?, canonical)
        },
        3 | 4 => return Err(buffa::DecodeError::InvalidWireType(wire as u32).into()),
        5 => Value::Fixed32(u32::from_le_bytes(
            take_exact(input, 4)?
                .try_into()
                .map_err(|_error| buffa::DecodeError::UnexpectedEof)?,
        )),
        _ => return Err(buffa::DecodeError::InvalidWireType(wire as u32).into()),
    };
    Ok(Some(Field {
        number,
        wire,
        value,
        canonical_key,
    }))
}

fn take_varint(input: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *input;
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
            *input = &original[consumed..];
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
    input: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], DecodeError> {
    if input.len() < length {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    let (selected, remaining) = input.split_at(length);
    *input = remaining;
    Ok(selected)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let max_buffa =
        usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| DecodeError::projection())?;
    if options.max_message_bytes > max_buffa {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: options.max_message_bytes,
            maximum: max_buffa,
        }));
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: source.len(),
            maximum: options.max_message_bytes,
        }));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::resource(WireResourceLimit::Nesting {
            observed: options.recursion_limit,
            maximum: MAX_RECURSION_LIMIT,
        }));
    }
    Ok(())
}

fn preflight_fill<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<BackgroundSnapshot<'source>, DecodeError> {
    budget.charge(source, options, depth)?;
    if source.is_empty() {
        return Ok(BackgroundSnapshot::None { raw: source });
    }
    let mut color = None;
    let mut gradient = None;
    let mut image = None;
    let mut seen = 0u16;
    let mut input = source;
    while let Some(field) = next_field(&mut input)? {
        match field.number {
            FILL_COLOR_FIELD => {
                unique_bit(&mut seen, 0, "TSD.FillArchive.color")?;
                let payload = field.bytes()?;
                color = Some(preflight_color(payload, options, budget, depth + 1)?);
            },
            FILL_GRADIENT_FIELD => {
                unique_bit(&mut seen, 1, "TSD.FillArchive.gradient")?;
                let payload = field.bytes()?;
                gradient = Some(preflight_gradient(payload, options, budget, depth + 1)?);
            },
            FILL_IMAGE_FIELD => {
                unique_bit(&mut seen, 2, "TSD.FillArchive.image")?;
                let payload = field.bytes()?;
                image = Some(preflight_image(payload, options, budget, depth + 1)?);
            },
            _ => {},
        }
    }
    match (color, gradient, image) {
        (Some(color), None, None) if is_supported_solid(color) => {
            Ok(BackgroundSnapshot::Solid { raw: source, color })
        },
        (None, Some(gradient), None) => {
            if gradient_is_supported(gradient, options, budget, depth + 1)? {
                Ok(BackgroundSnapshot::Gradient {
                    raw: source,
                    gradient,
                })
            } else {
                Ok(BackgroundSnapshot::Opaque { raw: source })
            }
        },
        (None, None, Some(image)) => Ok(BackgroundSnapshot::Image { raw: source, image }),
        _ => Ok(BackgroundSnapshot::Opaque { raw: source }),
    }
}

fn preflight_color(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<ColorSnapshot, DecodeError> {
    budget.charge(source, options, depth)?;
    let mut result = ColorSnapshot {
        model: 0,
        red: None,
        green: None,
        blue: None,
        alpha: None,
        cyan: None,
        magenta: None,
        yellow: None,
        black: None,
        white: None,
        rgb_space: None,
    };
    let mut seen = 0u16;
    let mut input = source;
    let mut model_seen = false;
    while let Some(field) = next_field(&mut input)? {
        match field.number {
            COLOR_MODEL_FIELD => {
                unique_bit(&mut seen, 0, "TSP.Color.model")?;
                result.model = field.int32()?;
                model_seen = true;
            },
            COLOR_RED_FIELD => result.red = Some(unique_float(&mut seen, 2, field, "TSP.Color.r")?),
            COLOR_GREEN_FIELD => {
                result.green = Some(unique_float(&mut seen, 3, field, "TSP.Color.g")?)
            },
            COLOR_BLUE_FIELD => {
                result.blue = Some(unique_float(&mut seen, 4, field, "TSP.Color.b")?)
            },
            COLOR_ALPHA_FIELD => {
                result.alpha = Some(unique_float(&mut seen, 5, field, "TSP.Color.a")?)
            },
            COLOR_CYAN_FIELD => {
                result.cyan = Some(unique_float(&mut seen, 6, field, "TSP.Color.c")?)
            },
            COLOR_MAGENTA_FIELD => {
                result.magenta = Some(unique_float(&mut seen, 7, field, "TSP.Color.m")?)
            },
            COLOR_YELLOW_FIELD => {
                result.yellow = Some(unique_float(&mut seen, 8, field, "TSP.Color.y")?)
            },
            COLOR_BLACK_FIELD => {
                result.black = Some(unique_float(&mut seen, 9, field, "TSP.Color.k")?)
            },
            COLOR_WHITE_FIELD => {
                result.white = Some(unique_float(&mut seen, 10, field, "TSP.Color.w")?)
            },
            COLOR_SPACE_FIELD => {
                unique_bit(&mut seen, 11, "TSP.Color.rgbspace")?;
                result.rgb_space = Some(field.int32()?);
            },
            _ => {},
        }
    }
    if !model_seen {
        return Err(DecodeError::missing_required("TSP.Color.model"));
    }
    // Native Keynote RGB conversion consumes all three channels and the
    // explicit color-space selector.  A message that claims RGB but omits one
    // of those selected fields is malformed rather than a future/opaque color
    // model; unsupported model values remain safely opaque below.
    if result.model == RGB_MODEL {
        if result.red.is_none() {
            return Err(DecodeError::missing_required("TSP.Color.r"));
        }
        if result.green.is_none() {
            return Err(DecodeError::missing_required("TSP.Color.g"));
        }
        if result.blue.is_none() {
            return Err(DecodeError::missing_required("TSP.Color.b"));
        }
        if result.rgb_space.is_none() {
            return Err(DecodeError::missing_required("TSP.Color.rgbspace"));
        }
    }
    Ok(result)
}

fn is_supported_solid(color: ColorSnapshot) -> bool {
    color.model == RGB_MODEL
        && color.red.is_some()
        && color.green.is_some()
        && color.blue.is_some()
        && color
            .rgb_space
            .is_some_and(|space| matches!(space, SRGB_SPACE | P3_SPACE))
        && color.cyan.is_none()
        && color.magenta.is_none()
        && color.yellow.is_none()
        && color.black.is_none()
        && color.white.is_none()
}

fn preflight_gradient<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<GradientSnapshot<'source>, DecodeError> {
    budget.charge(source, options, depth)?;
    let mut gradient_type = None;
    let mut opacity = None;
    let mut advanced = None;
    let mut angle_radians = None;
    let mut transform = None;
    let mut seen = 0u16;
    let mut input = source;
    while let Some(field) = next_field(&mut input)? {
        match field.number {
            GRADIENT_TYPE_FIELD => {
                unique_bit(&mut seen, 0, "TSD.GradientArchive.type")?;
                gradient_type = Some(field.int32()?);
            },
            GRADIENT_STOP_FIELD => {
                // Repeated by schema; every element is independently bounded
                // and checked for required nested color/point values.
                let payload = field.bytes()?;
                parse_gradient_stop_with_budget(payload, options, budget, depth + 1)?;
            },
            GRADIENT_OPACITY_FIELD => {
                opacity = Some(unique_float(
                    &mut seen,
                    1,
                    field,
                    "TSD.GradientArchive.opacity",
                )?);
            },
            GRADIENT_ADVANCED_FIELD => {
                unique_bit(&mut seen, 2, "TSD.GradientArchive.advancedGradient")?;
                advanced = Some(field.bool()?);
            },
            GRADIENT_ANGLE_FIELD => {
                unique_bit(&mut seen, 3, "TSD.GradientArchive.anglegradient")?;
                let payload = field.bytes()?;
                angle_radians = preflight_angle(payload, options, budget, depth + 1)?;
            },
            GRADIENT_TRANSFORM_FIELD => {
                unique_bit(&mut seen, 4, "TSD.GradientArchive.transformgradient")?;
                let payload = field.bytes()?;
                preflight_transform(payload, options, budget, depth + 1)?;
                transform = Some(payload);
            },
            _ => {},
        }
    }
    Ok(GradientSnapshot {
        raw: source,
        gradient_type,
        opacity,
        advanced,
        angle_radians,
        transform,
    })
}

fn gradient_is_supported(
    gradient: GradientSnapshot<'_>,
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<bool, DecodeError> {
    if !matches!(gradient.gradient_type, Some(0 | 1))
        || gradient.opacity.is_none()
        || gradient.advanced.is_none()
        || gradient.angle_radians.is_none()
        || gradient.transform.is_some()
    {
        return Ok(false);
    }
    charge_gradient_iterator_pass(gradient.raw, options, budget, depth)?;
    let mut count = 0usize;
    let mut previous_fraction = None;
    let mut simple_shape = true;
    for result in gradient.stops(options) {
        let stop = result?;
        count += 1;
        let (Some(color), Some(fraction), Some(inflection)) =
            (stop.color, stop.fraction, stop.inflection)
        else {
            return Ok(false);
        };
        if previous_fraction.is_some_and(|previous| fraction < previous) {
            return Ok(false);
        }
        previous_fraction = Some(fraction);
        if inflection.to_bits() != 0.5f32.to_bits() {
            simple_shape = false;
        }
        if !is_supported_solid(color) {
            return Ok(false);
        }
    }
    let advanced = gradient.advanced.expect("checked above");
    if !advanced && (gradient.gradient_type != Some(0) || count != 2 || !simple_shape) {
        return Ok(false);
    }
    Ok(count >= 2)
}

fn charge_gradient_iterator_pass(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.charge(source, options, depth)?;
    let mut input = source;
    while let Some(field) = next_field(&mut input)? {
        if field.number != GRADIENT_STOP_FIELD {
            continue;
        }
        let stop = field.bytes()?;
        budget.charge(stop, options, depth + 1)?;
        let mut stop_input = stop;
        while let Some(stop_field) = next_field(&mut stop_input)? {
            if stop_field.number == STOP_COLOR_FIELD {
                budget.charge(stop_field.bytes()?, options, depth + 2)?;
            }
        }
    }
    Ok(())
}

fn parse_gradient_stop(
    source: &[u8],
    options: DecodeOptions,
) -> Result<GradientStopSnapshot, DecodeError> {
    // The iterator is used after the strict pass, so this second read does not
    // charge or allocate; it still uses the same canonical parser to ensure a
    // caller cannot manufacture a snapshot from an unvalidated raw slice.
    let mut input = source;
    let mut color = None;
    let mut fraction = None;
    let mut inflection = None;
    let mut seen = 0u16;
    while let Some(field) = next_field(&mut input)? {
        match field.number {
            STOP_COLOR_FIELD => {
                unique_bit(&mut seen, 0, "TSD.GradientStop.color")?;
                color = Some(preflight_color(
                    field.bytes()?,
                    options,
                    &mut Budget::new(options),
                    0,
                )?);
            },
            STOP_FRACTION_FIELD => {
                unique_bit(&mut seen, 1, "TSD.GradientStop.fraction")?;
                fraction = Some(bounded_float(field, "gradient stop fraction")?);
            },
            STOP_INFLECTION_FIELD => {
                unique_bit(&mut seen, 2, "TSD.GradientStop.inflection")?;
                inflection = Some(bounded_float(field, "gradient stop inflection")?);
            },
            _ => {},
        }
    }
    Ok(GradientStopSnapshot {
        color,
        fraction,
        inflection,
    })
}

fn parse_gradient_stop_with_budget(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.charge(source, options, depth)?;
    let mut input = source;
    let mut seen = 0u16;
    while let Some(field) = next_field(&mut input)? {
        match field.number {
            STOP_COLOR_FIELD => {
                unique_bit(&mut seen, 0, "TSD.GradientStop.color")?;
                preflight_color(field.bytes()?, options, budget, depth + 1)?;
            },
            STOP_FRACTION_FIELD => {
                unique_bit(&mut seen, 1, "TSD.GradientStop.fraction")?;
                bounded_float(field, "gradient stop fraction")?;
            },
            STOP_INFLECTION_FIELD => {
                unique_bit(&mut seen, 2, "TSD.GradientStop.inflection")?;
                bounded_float(field, "gradient stop inflection")?;
            },
            _ => {},
        }
    }
    Ok(())
}

fn preflight_angle(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<f32>, DecodeError> {
    budget.charge(source, options, depth)?;
    let mut input = source;
    let mut result = None;
    let mut seen = false;
    while let Some(field) = next_field(&mut input)? {
        if field.number == ANGLE_RADIANS_FIELD {
            if seen {
                return Err(DecodeError::duplicate_singular(
                    "TSD.AngleGradientArchive.gradientangle",
                ));
            }
            seen = true;
            result = Some(angle_float(field)?);
        }
    }
    Ok(result)
}

fn preflight_transform(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.charge(source, options, depth)?;
    let mut input = source;
    let mut seen = 0u16;
    while let Some(field) = next_field(&mut input)? {
        match field.number {
            TRANSFORM_START_FIELD | TRANSFORM_END_FIELD => {
                unique_bit(
                    &mut seen,
                    field.number as u8 - 1,
                    "TSD.TransformGradientArchive.point",
                )?;
                preflight_point(field.bytes()?, options, budget, depth + 1)?;
            },
            TRANSFORM_SIZE_FIELD => {
                unique_bit(&mut seen, 2, "TSD.TransformGradientArchive.baseNaturalSize")?;
                preflight_size(field.bytes()?, options, budget, depth + 1)?;
            },
            _ => {},
        }
    }
    Ok(())
}

fn preflight_image<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<ImageSnapshot<'source>, DecodeError> {
    budget.charge(source, options, depth)?;
    let mut result = ImageSnapshot {
        raw: source,
        database_image_identifier: None,
        technique: None,
        tint: None,
        fill_size: None,
        database_original_identifier: None,
        image_identifier: None,
        original_image_identifier: None,
        interprets_untagged_image_data_as_generic: None,
        reference_color: None,
    };
    let mut seen = 0u16;
    let mut input = source;
    while let Some(field) = next_field(&mut input)? {
        match field.number {
            IMAGE_DATABASE_DATA_FIELD => {
                unique_bit(&mut seen, 0, "TSD.ImageFillArchive.database_imagedata")?;
                result.database_image_identifier = Some(preflight_reference(
                    field.bytes()?,
                    options,
                    budget,
                    depth + 1,
                    "TSP.Reference.identifier",
                )?);
            },
            IMAGE_TECHNIQUE_FIELD => {
                unique_bit(&mut seen, 1, "TSD.ImageFillArchive.technique")?;
                result.technique = Some(field.int32()?);
            },
            IMAGE_TINT_FIELD => {
                unique_bit(&mut seen, 2, "TSD.ImageFillArchive.tint")?;
                result.tint = Some(preflight_color(field.bytes()?, options, budget, depth + 1)?);
            },
            IMAGE_SIZE_FIELD => {
                unique_bit(&mut seen, 3, "TSD.ImageFillArchive.fillsize")?;
                result.fill_size =
                    Some(preflight_size(field.bytes()?, options, budget, depth + 1)?);
            },
            IMAGE_DATABASE_ORIGINAL_FIELD => {
                unique_bit(
                    &mut seen,
                    4,
                    "TSD.ImageFillArchive.database_originalimagedata",
                )?;
                result.database_original_identifier = Some(preflight_reference(
                    field.bytes()?,
                    options,
                    budget,
                    depth + 1,
                    "TSP.Reference.identifier",
                )?);
            },
            IMAGE_DATA_FIELD => {
                unique_bit(&mut seen, 5, "TSD.ImageFillArchive.imagedata")?;
                result.image_identifier = Some(preflight_data_reference(
                    field.bytes()?,
                    options,
                    budget,
                    depth + 1,
                )?);
            },
            IMAGE_ORIGINAL_DATA_FIELD => {
                unique_bit(&mut seen, 6, "TSD.ImageFillArchive.originalimagedata")?;
                result.original_image_identifier = Some(preflight_data_reference(
                    field.bytes()?,
                    options,
                    budget,
                    depth + 1,
                )?);
            },
            IMAGE_UNTAGGED_FIELD => {
                unique_bit(
                    &mut seen,
                    7,
                    "TSD.ImageFillArchive.interpretsUntaggedImageDataAsGeneric",
                )?;
                result.interprets_untagged_image_data_as_generic = Some(field.bool()?);
            },
            IMAGE_REFERENCE_COLOR_FIELD => {
                unique_bit(&mut seen, 8, "TSD.ImageFillArchive.referencecolor")?;
                result.reference_color =
                    Some(preflight_color(field.bytes()?, options, budget, depth + 1)?);
            },
            _ => {},
        }
    }
    Ok(result)
}

fn preflight_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
    context: &'static str,
) -> Result<u64, DecodeError> {
    budget.charge(source, options, depth)?;
    let mut input = source;
    let mut identifier = None;
    let mut seen = 0u8;
    while let Some(field) = next_field(&mut input)? {
        match field.number {
            RESOURCE_IDENTIFIER_FIELD => {
                unique_bit_u8(&mut seen, 0, context)?;
                let value = field.varint()?;
                if value == 0 {
                    return Err(DecodeError::noncanonical("resource identifier is zero"));
                }
                identifier = Some(value);
            },
            2 => {
                // Deprecated Reference fields are optional but remain strict
                // scalar values when present.
                unique_bit_u8(&mut seen, 1, "TSP.Reference.deprecated_type")?;
                field.int32()?;
            },
            3 => {
                unique_bit_u8(&mut seen, 2, "TSP.Reference.deprecated_is_external")?;
                field.bool()?;
            },
            _ => {},
        }
    }
    identifier.ok_or_else(|| DecodeError::missing_required(context))
}

fn preflight_data_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<u64, DecodeError> {
    budget.charge(source, options, depth)?;
    let mut input = source;
    let mut identifier = None;
    while let Some(field) = next_field(&mut input)? {
        if field.number == RESOURCE_IDENTIFIER_FIELD {
            if identifier.is_some() {
                return Err(DecodeError::duplicate_singular(
                    "TSP.DataReference.identifier",
                ));
            }
            let value = field.varint()?;
            if value == 0 {
                return Err(DecodeError::noncanonical("resource identifier is zero"));
            }
            identifier = Some(value);
        }
    }
    identifier.ok_or_else(|| DecodeError::missing_required("TSP.DataReference.identifier"))
}

fn preflight_point(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.charge(source, options, depth)?;
    let mut input = source;
    let mut seen = 0u16;
    while let Some(field) = next_field(&mut input)? {
        match field.number {
            POINT_X_FIELD => {
                unique_bit(&mut seen, 0, "TSP.Point.x")?;
                field.float()?;
            },
            POINT_Y_FIELD => {
                unique_bit(&mut seen, 1, "TSP.Point.y")?;
                field.float()?;
            },
            _ => {},
        }
    }
    if seen & 0b11 != 0b11 {
        return Err(DecodeError::missing_required("TSP.Point.x/y"));
    }
    Ok(())
}

fn preflight_size(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<SizeSnapshot, DecodeError> {
    budget.charge(source, options, depth)?;
    let mut input = source;
    let mut seen = 0u16;
    let mut width = None;
    let mut height = None;
    while let Some(field) = next_field(&mut input)? {
        match field.number {
            SIZE_WIDTH_FIELD => {
                unique_bit(&mut seen, 0, "TSP.Size.width")?;
                width = Some(field.float()?);
            },
            SIZE_HEIGHT_FIELD => {
                unique_bit(&mut seen, 1, "TSP.Size.height")?;
                height = Some(field.float()?);
            },
            _ => {},
        }
    }
    Ok(SizeSnapshot {
        width: width.ok_or_else(|| DecodeError::missing_required("TSP.Size.width"))?,
        height: height.ok_or_else(|| DecodeError::missing_required("TSP.Size.height"))?,
    })
}

fn cross_check_projection(
    view: &projection::KeynoteSlideBackgroundArchiveLazyView<'_>,
    strict: &BackgroundSnapshot<'_>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge(strict.raw(), options, 0)?;
    let projected_color = view.color.get()?.map(|color| ColorSnapshot {
        model: color.model,
        red: color.r,
        green: color.g,
        blue: color.b,
        alpha: color.a,
        cyan: color.c,
        magenta: color.m,
        yellow: color.y,
        black: color.k,
        white: color.w,
        rgb_space: color.rgbspace,
    });
    let projected_gradient = view.gradient;
    let projected_image = view.image;
    let source = strict.raw();
    if let Some(color) = find_root_payload(source, FILL_COLOR_FIELD)? {
        budget.charge(color, options, 1)?;
    }
    let strict_color = find_root_payload(source, FILL_COLOR_FIELD)?
        .map(parse_color_snapshot_for_projection)
        .transpose()?;
    let strict_gradient = find_root_payload(source, FILL_GRADIENT_FIELD)?;
    let strict_image = find_root_payload(source, FILL_IMAGE_FIELD)?;
    if !same_color(projected_color, strict_color)
        || projected_gradient != strict_gradient
        || projected_image != strict_image
    {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn parse_color_snapshot_for_projection(source: &[u8]) -> Result<ColorSnapshot, DecodeError> {
    let options = DecodeOptions::new(source.len().max(1), MAX_RECURSION_LIMIT)
        .with_resource_limits(usize::MAX, usize::MAX);
    let mut budget = Budget::new(options);
    preflight_color(source, options, &mut budget, 0)
}

fn same_color(left: Option<ColorSnapshot>, right: Option<ColorSnapshot>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => {
            left.model == right.model
                && left.red.map(f32::to_bits) == right.red.map(f32::to_bits)
                && left.green.map(f32::to_bits) == right.green.map(f32::to_bits)
                && left.blue.map(f32::to_bits) == right.blue.map(f32::to_bits)
                && left.alpha.map(f32::to_bits) == right.alpha.map(f32::to_bits)
                && left.cyan.map(f32::to_bits) == right.cyan.map(f32::to_bits)
                && left.magenta.map(f32::to_bits) == right.magenta.map(f32::to_bits)
                && left.yellow.map(f32::to_bits) == right.yellow.map(f32::to_bits)
                && left.black.map(f32::to_bits) == right.black.map(f32::to_bits)
                && left.white.map(f32::to_bits) == right.white.map(f32::to_bits)
                && left.rgb_space == right.rgb_space
        },
        (None, Some(_)) | (Some(_), None) => false,
    }
}

fn find_root_payload(source: &[u8], number: u32) -> Result<Option<&[u8]>, DecodeError> {
    let mut input = source;
    while let Some(field) = next_field(&mut input)? {
        if field.number == number {
            return Ok(Some(field.bytes()?));
        }
    }
    Ok(None)
}

fn unique_bit(seen: &mut u16, bit: u8, field: &'static str) -> Result<(), DecodeError> {
    let mask = 1u16
        .checked_shl(u32::from(bit))
        .ok_or_else(DecodeError::projection)?;
    if *seen & mask != 0 {
        return Err(DecodeError::duplicate_singular(field));
    }
    *seen |= mask;
    Ok(())
}

fn unique_bit_u8(seen: &mut u8, bit: u8, field: &'static str) -> Result<(), DecodeError> {
    let mask = 1u8
        .checked_shl(u32::from(bit))
        .ok_or_else(DecodeError::projection)?;
    if *seen & mask != 0 {
        return Err(DecodeError::duplicate_singular(field));
    }
    *seen |= mask;
    Ok(())
}

fn unique_float(
    seen: &mut u16,
    bit: u8,
    field: Field<'_>,
    name: &'static str,
) -> Result<f32, DecodeError> {
    unique_bit(seen, bit, name)?;
    let value = field.float()?;
    if !(0.0..=1.0).contains(&value) {
        return Err(DecodeError::noncanonical("color component outside [0,1]"));
    }
    Ok(value)
}

fn bounded_float(field: Field<'_>, context: &'static str) -> Result<f32, DecodeError> {
    let value = field.float()?;
    if !(0.0..=1.0).contains(&value) {
        return Err(DecodeError::noncanonical(context));
    }
    Ok(value)
}

fn angle_float(field: Field<'_>) -> Result<f32, DecodeError> {
    let value = field.float()?;
    if !(0.0..std::f32::consts::TAU).contains(&value) {
        return Err(DecodeError::noncanonical("gradient angle outside [0,2π)"));
    }
    Ok(value)
}

fn encode_write<'source>(
    source: &'source [u8],
    snapshot: BackgroundSnapshot<'source>,
    write: BackgroundWrite<'source>,
    max_message_bytes: usize,
) -> Result<Vec<u8>, DecodeError> {
    match write {
        BackgroundWrite::Clear => Ok(Vec::new()),
        BackgroundWrite::Raw(payload) => {
            let mut output = try_vec_bounded(payload.len(), max_message_bytes)?;
            output.extend_from_slice(payload);
            Ok(output)
        },
        BackgroundWrite::Gradient(gradient) => {
            rewrite_gradient_fill(source, gradient, max_message_bytes)
        },
        BackgroundWrite::Solid(color) => {
            rewrite_solid_fill(source, snapshot, color, max_message_bytes)
        },
    }
}

fn measure_write_output_len(
    source: &[u8],
    write: BackgroundWrite<'_>,
) -> Result<OutputMeasure, DecodeError> {
    match write {
        BackgroundWrite::Clear => Ok(OutputMeasure {
            bytes: 0,
            scan_bytes: 0,
        }),
        BackgroundWrite::Raw(payload) => Ok(OutputMeasure {
            bytes: payload.len(),
            scan_bytes: 0,
        }),
        BackgroundWrite::Solid(color) => measure_solid_output_len(source, color),
        BackgroundWrite::Gradient(gradient) => measure_gradient_output_len(source, gradient),
    }
}

fn count_rewrite_allocations(
    source: &[u8],
    write: BackgroundWrite<'_>,
) -> Result<usize, DecodeError> {
    match write {
        BackgroundWrite::Clear => Ok(0),
        BackgroundWrite::Raw(_) => Ok(1),
        BackgroundWrite::Solid(_) => {
            let has_color = find_root_payload(source, FILL_COLOR_FIELD)?.is_some();
            Ok(if has_color { 3 } else { 2 })
        },
        BackgroundWrite::Gradient(gradient) => {
            let has_gradient = find_root_payload(source, FILL_GRADIENT_FIELD)?.is_some();
            let base = gradient
                .stops
                .len()
                .checked_mul(2)
                .and_then(|value| value.checked_add(3))
                .ok_or_else(DecodeError::projection)?;
            if has_gradient {
                base.checked_add(2).ok_or_else(DecodeError::projection)
            } else {
                Ok(base)
            }
        },
    }
}

fn measure_solid_output_len(
    source: &[u8],
    color: ColorWrite,
) -> Result<OutputMeasure, DecodeError> {
    validate_color_write(color)?;
    let replacement_len = encoded_color_len();
    let mut output_len = 0usize;
    let mut input = source;
    let mut found_color = false;
    let mut scan_bytes = source.len();
    while !input.is_empty() {
        let before = input.len();
        let field = next_field(&mut input)?.ok_or_else(DecodeError::projection)?;
        let consumed = before - input.len();
        let raw_len = consumed;
        match field.number {
            FILL_COLOR_FIELD => {
                if found_color {
                    return Err(DecodeError::duplicate_singular("TSD.FillArchive.color"));
                }
                found_color = true;
                let payload = field.bytes()?;
                let (rewritten, color_scan_bytes) = measure_color_payload_len(payload, color)?;
                scan_bytes = scan_bytes
                    .checked_add(color_scan_bytes)
                    .ok_or_else(DecodeError::projection)?;
                output_len = output_len
                    .checked_add(length_field_len(FILL_COLOR_FIELD, rewritten))
                    .ok_or_else(DecodeError::projection)?;
            },
            FILL_GRADIENT_FIELD | FILL_IMAGE_FIELD => {},
            _ => {
                output_len = output_len
                    .checked_add(raw_len)
                    .ok_or_else(DecodeError::projection)?;
            },
        }
    }
    if !found_color {
        output_len = output_len
            .checked_add(length_field_len(FILL_COLOR_FIELD, replacement_len))
            .ok_or_else(DecodeError::projection)?;
    }
    Ok(OutputMeasure {
        bytes: output_len,
        scan_bytes,
    })
}

fn measure_color_payload_len(
    source: &[u8],
    color: ColorWrite,
) -> Result<(usize, usize), DecodeError> {
    let mut output_len = 0usize;
    let mut input = source;
    let mut seen = 0u16;
    while !input.is_empty() {
        let before = input.len();
        let field = next_field(&mut input)?.ok_or_else(DecodeError::projection)?;
        let consumed = before - input.len();
        match field.number {
            COLOR_MODEL_FIELD => {
                unique_bit(&mut seen, 0, "TSP.Color.model")?;
                field.int32()?;
                output_len = output_len
                    .checked_add(varint_field_len(COLOR_MODEL_FIELD, RGB_MODEL as u64))
                    .ok_or_else(DecodeError::projection)?;
            },
            COLOR_RED_FIELD => {
                unique_bit(&mut seen, 1, "TSP.Color.r")?;
                field.float()?;
                output_len = output_len
                    .checked_add(fixed32_field_len(COLOR_RED_FIELD))
                    .ok_or_else(DecodeError::projection)?;
            },
            COLOR_GREEN_FIELD => {
                unique_bit(&mut seen, 2, "TSP.Color.g")?;
                field.float()?;
                output_len = output_len
                    .checked_add(fixed32_field_len(COLOR_GREEN_FIELD))
                    .ok_or_else(DecodeError::projection)?;
            },
            COLOR_BLUE_FIELD => {
                unique_bit(&mut seen, 3, "TSP.Color.b")?;
                field.float()?;
                output_len = output_len
                    .checked_add(fixed32_field_len(COLOR_BLUE_FIELD))
                    .ok_or_else(DecodeError::projection)?;
            },
            COLOR_ALPHA_FIELD => {
                unique_bit(&mut seen, 4, "TSP.Color.a")?;
                field.float()?;
                output_len = output_len
                    .checked_add(fixed32_field_len(COLOR_ALPHA_FIELD))
                    .ok_or_else(DecodeError::projection)?;
            },
            COLOR_SPACE_FIELD => {
                unique_bit(&mut seen, 5, "TSP.Color.rgbspace")?;
                field.int32()?;
                output_len = output_len
                    .checked_add(varint_field_len(COLOR_SPACE_FIELD, rgb_space_value(color)))
                    .ok_or_else(DecodeError::projection)?;
            },
            COLOR_CYAN_FIELD | COLOR_MAGENTA_FIELD | COLOR_YELLOW_FIELD | COLOR_BLACK_FIELD
            | COLOR_WHITE_FIELD => {},
            _ => {
                output_len = output_len
                    .checked_add(consumed)
                    .ok_or_else(DecodeError::projection)?;
            },
        }
    }
    let missing = [
        (0, varint_field_len(COLOR_MODEL_FIELD, RGB_MODEL as u64)),
        (1, fixed32_field_len(COLOR_RED_FIELD)),
        (2, fixed32_field_len(COLOR_GREEN_FIELD)),
        (3, fixed32_field_len(COLOR_BLUE_FIELD)),
        (4, fixed32_field_len(COLOR_ALPHA_FIELD)),
        (
            5,
            varint_field_len(COLOR_SPACE_FIELD, rgb_space_value(color)),
        ),
    ];
    for (bit, length) in missing {
        if seen & (1 << bit) == 0 {
            output_len = output_len
                .checked_add(length)
                .ok_or_else(DecodeError::projection)?;
        }
    }
    Ok((output_len, source.len()))
}

fn measure_gradient_output_len(
    source: &[u8],
    gradient: GradientWrite<'_>,
) -> Result<OutputMeasure, DecodeError> {
    let replacement_len = encoded_gradient_payload_len(gradient)?;
    let mut output_len = 0usize;
    let mut input = source;
    let mut found_gradient = false;
    let mut scan_bytes = source.len();
    while !input.is_empty() {
        let before = input.len();
        let field = next_field(&mut input)?.ok_or_else(DecodeError::projection)?;
        let consumed = before - input.len();
        match field.number {
            FILL_COLOR_FIELD | FILL_IMAGE_FIELD => {},
            FILL_GRADIENT_FIELD => {
                if found_gradient {
                    return Err(DecodeError::duplicate_singular("TSD.FillArchive.gradient"));
                }
                found_gradient = true;
                let payload = field.bytes()?;
                let (unknown_len, nested_scan_bytes) = measure_unknown_gradient_fields(payload)?;
                scan_bytes = scan_bytes
                    .checked_add(nested_scan_bytes)
                    .ok_or_else(DecodeError::projection)?;
                let rewritten_len = replacement_len
                    .checked_add(unknown_len)
                    .ok_or_else(DecodeError::projection)?;
                output_len = output_len
                    .checked_add(length_field_len(FILL_GRADIENT_FIELD, rewritten_len))
                    .ok_or_else(DecodeError::projection)?;
            },
            _ => {
                output_len = output_len
                    .checked_add(consumed)
                    .ok_or_else(DecodeError::projection)?;
            },
        }
    }
    if !found_gradient {
        output_len = output_len
            .checked_add(length_field_len(FILL_GRADIENT_FIELD, replacement_len))
            .ok_or_else(DecodeError::projection)?;
    }
    Ok(OutputMeasure {
        bytes: output_len,
        scan_bytes,
    })
}

fn measure_unknown_gradient_fields(source: &[u8]) -> Result<(usize, usize), DecodeError> {
    let mut input = source;
    let mut length = 0usize;
    while !input.is_empty() {
        let before = input.len();
        let field = next_field(&mut input)?.ok_or_else(DecodeError::projection)?;
        let consumed = before - input.len();
        if field.number > GRADIENT_TRANSFORM_FIELD {
            length = length
                .checked_add(consumed)
                .ok_or_else(DecodeError::projection)?;
        }
    }
    Ok((length, source.len()))
}

fn encoded_color_len() -> usize {
    varint_field_len(COLOR_MODEL_FIELD, RGB_MODEL as u64)
        + fixed32_field_len(COLOR_RED_FIELD)
        + fixed32_field_len(COLOR_GREEN_FIELD)
        + fixed32_field_len(COLOR_BLUE_FIELD)
        + fixed32_field_len(COLOR_ALPHA_FIELD)
        + varint_field_len(COLOR_SPACE_FIELD, SRGB_SPACE as u64)
}

fn encoded_gradient_payload_len(gradient: GradientWrite<'_>) -> Result<usize, DecodeError> {
    validate_gradient_write(gradient)?;
    let stop_color_len = encoded_color_len();
    let stop_len = length_field_len(STOP_COLOR_FIELD, stop_color_len)
        + fixed32_field_len(STOP_FRACTION_FIELD)
        + fixed32_field_len(STOP_INFLECTION_FIELD);
    let angle_len = fixed32_field_len(ANGLE_RADIANS_FIELD);
    let mut length = varint_field_len(
        GRADIENT_TYPE_FIELD,
        match gradient.kind {
            GradientKind::Linear => 0,
            GradientKind::Radial => 1,
        },
    );
    length = length
        .checked_add(
            gradient
                .stops
                .len()
                .checked_mul(length_field_len(GRADIENT_STOP_FIELD, stop_len))
                .ok_or_else(DecodeError::projection)?,
        )
        .and_then(|value| value.checked_add(fixed32_field_len(GRADIENT_OPACITY_FIELD)))
        .and_then(|value| {
            value.checked_add(varint_field_len(
                GRADIENT_ADVANCED_FIELD,
                u64::from(gradient.advanced),
            ))
        })
        .and_then(|value| value.checked_add(length_field_len(GRADIENT_ANGLE_FIELD, angle_len)))
        .ok_or_else(DecodeError::projection)?;
    Ok(length)
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn varint_field_len(number: u32, value: u64) -> usize {
    varint_len(u64::from(number) << 3) + varint_len(value)
}

fn fixed32_field_len(number: u32) -> usize {
    varint_len((u64::from(number) << 3) | 5) + 4
}

fn length_field_len(number: u32, payload_len: usize) -> usize {
    varint_len((u64::from(number) << 3) | 2) + varint_len(payload_len as u64) + payload_len
}

fn rgb_space_value(color: ColorWrite) -> u64 {
    match color.rgb_space {
        RgbSpace::Srgb => SRGB_SPACE as u64,
        RgbSpace::DisplayP3 => P3_SPACE as u64,
    }
}

fn ensure_write_matches(
    write: BackgroundWrite<'_>,
    snapshot: &BackgroundSnapshot<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    match write {
        BackgroundWrite::Clear => {
            if !matches!(snapshot, BackgroundSnapshot::None { .. }) {
                return Err(DecodeError::projection());
            }
        },
        BackgroundWrite::Solid(color) => {
            let BackgroundSnapshot::Solid { color: actual, .. } = snapshot else {
                return Err(DecodeError::projection());
            };
            if actual.model != RGB_MODEL
                || actual.red.map(f32::to_bits) != Some(color.red.to_bits())
                || actual.green.map(f32::to_bits) != Some(color.green.to_bits())
                || actual.blue.map(f32::to_bits) != Some(color.blue.to_bits())
                || actual.alpha.unwrap_or(1.0).to_bits() != color.alpha.to_bits()
                || actual.rgb_space
                    != Some(match color.rgb_space {
                        RgbSpace::Srgb => SRGB_SPACE,
                        RgbSpace::DisplayP3 => P3_SPACE,
                    })
            {
                return Err(DecodeError::projection());
            }
        },
        BackgroundWrite::Gradient(gradient) => {
            if !matches!(snapshot, BackgroundSnapshot::Gradient { .. }) {
                return Err(DecodeError::projection());
            }
            let BackgroundSnapshot::Gradient {
                gradient: actual, ..
            } = snapshot
            else {
                unreachable!("matches checked above")
            };
            let expected_type = match gradient.kind {
                GradientKind::Linear => 0,
                GradientKind::Radial => 1,
            };
            if actual.gradient_type != Some(expected_type)
                || actual.opacity.map(f32::to_bits) != Some(gradient.opacity.to_bits())
                || actual.advanced != Some(gradient.advanced)
                || actual.angle_radians.map(f32::to_bits) != Some(gradient.angle_radians.to_bits())
            {
                return Err(DecodeError::projection());
            }
            let mut actual_stops = actual.stops(options);
            for expected in gradient.stops {
                let Some(actual_stop) = actual_stops.next() else {
                    return Err(DecodeError::projection());
                };
                let actual_stop = actual_stop?;
                if !same_stop(actual_stop, *expected) {
                    return Err(DecodeError::projection());
                }
            }
            if let Some(extra) = actual_stops.next() {
                let _ = extra?;
                return Err(DecodeError::projection());
            }
        },
        BackgroundWrite::Raw(payload) => {
            if snapshot.raw() != payload {
                return Err(DecodeError::projection());
            }
        },
    }
    Ok(())
}

fn same_stop(actual: GradientStopSnapshot, expected: GradientStopWrite) -> bool {
    actual.fraction.map(f32::to_bits) == Some(expected.fraction.to_bits())
        && actual.inflection.map(f32::to_bits) == Some(expected.inflection.to_bits())
        && actual
            .color
            .is_some_and(|actual| same_color_value(actual, expected.color))
}

fn same_color_value(actual: ColorSnapshot, expected: ColorWrite) -> bool {
    actual.model == RGB_MODEL
        && actual.red.map(f32::to_bits) == Some(expected.red.to_bits())
        && actual.green.map(f32::to_bits) == Some(expected.green.to_bits())
        && actual.blue.map(f32::to_bits) == Some(expected.blue.to_bits())
        && actual.alpha.unwrap_or(1.0).to_bits() == expected.alpha.to_bits()
        && actual.rgb_space
            == Some(match expected.rgb_space {
                RgbSpace::Srgb => SRGB_SPACE,
                RgbSpace::DisplayP3 => P3_SPACE,
            })
        && actual.cyan.is_none()
        && actual.magenta.is_none()
        && actual.yellow.is_none()
        && actual.black.is_none()
        && actual.white.is_none()
}

fn validate_color_write(color: ColorWrite) -> Result<(), DecodeError> {
    for value in [color.red, color.green, color.blue, color.alpha] {
        if !value.is_finite() || !(0.0..=1.0).contains(&value) {
            return Err(DecodeError::noncanonical(
                "solid color component outside [0,1]",
            ));
        }
    }
    Ok(())
}

fn validate_gradient_write(gradient: GradientWrite<'_>) -> Result<(), DecodeError> {
    if !gradient.opacity.is_finite()
        || !(0.0..=1.0).contains(&gradient.opacity)
        || !gradient.angle_radians.is_finite()
        || !(0.0..std::f32::consts::TAU).contains(&gradient.angle_radians)
        || gradient.stops.len() < 2
    {
        return Err(DecodeError::noncanonical("invalid gradient scalar"));
    }
    if !gradient.advanced
        && (gradient.kind != GradientKind::Linear
            || gradient.stops.len() != 2
            || gradient
                .stops
                .iter()
                .any(|stop| stop.inflection.to_bits() != 0.5f32.to_bits()))
    {
        return Err(DecodeError::noncanonical("invalid simple gradient shape"));
    }
    let mut previous = 0.0f32;
    for (index, stop) in gradient.stops.iter().enumerate() {
        validate_color_write(stop.color)?;
        if !stop.fraction.is_finite()
            || !(0.0..=1.0).contains(&stop.fraction)
            || !stop.inflection.is_finite()
            || !(0.0..=1.0).contains(&stop.inflection)
            || (index != 0 && stop.fraction < previous)
        {
            return Err(DecodeError::noncanonical("invalid gradient stop scalar"));
        }
        previous = stop.fraction;
    }
    Ok(())
}

fn encode_gradient_payload(
    gradient: GradientWrite<'_>,
    max_message_bytes: usize,
) -> Result<Vec<u8>, DecodeError> {
    validate_gradient_write(gradient)?;
    let gradient_capacity = gradient.stops.len().saturating_mul(64).saturating_add(64);
    let mut gradient_payload = try_vec_bounded(gradient_capacity, max_message_bytes)?;
    append_varint_field(
        &mut gradient_payload,
        GRADIENT_TYPE_FIELD,
        match gradient.kind {
            GradientKind::Linear => 0,
            GradientKind::Radial => 1,
        },
    );
    for stop in gradient.stops {
        let mut stop_payload = try_vec_bounded(64, max_message_bytes)?;
        let encoded_color = encode_color(stop.color, max_message_bytes)?;
        append_length_field(&mut stop_payload, STOP_COLOR_FIELD, &encoded_color);
        append_fixed32_field(&mut stop_payload, STOP_FRACTION_FIELD, stop.fraction);
        append_fixed32_field(&mut stop_payload, STOP_INFLECTION_FIELD, stop.inflection);
        append_length_field(&mut gradient_payload, GRADIENT_STOP_FIELD, &stop_payload);
    }
    append_fixed32_field(
        &mut gradient_payload,
        GRADIENT_OPACITY_FIELD,
        gradient.opacity,
    );
    append_varint_field(
        &mut gradient_payload,
        GRADIENT_ADVANCED_FIELD,
        u64::from(gradient.advanced),
    );
    let mut angle_payload = try_vec_bounded(16, max_message_bytes)?;
    append_fixed32_field(
        &mut angle_payload,
        ANGLE_RADIANS_FIELD,
        gradient.angle_radians,
    );
    append_length_field(&mut gradient_payload, GRADIENT_ANGLE_FIELD, &angle_payload);
    Ok(gradient_payload)
}

fn encode_color(color: ColorWrite, max_message_bytes: usize) -> Result<Vec<u8>, DecodeError> {
    let mut payload = try_vec_bounded(64, max_message_bytes)?;
    append_varint_field(&mut payload, COLOR_MODEL_FIELD, RGB_MODEL as u64);
    append_fixed32_field(&mut payload, COLOR_RED_FIELD, color.red);
    append_fixed32_field(&mut payload, COLOR_GREEN_FIELD, color.green);
    append_fixed32_field(&mut payload, COLOR_BLUE_FIELD, color.blue);
    append_fixed32_field(&mut payload, COLOR_ALPHA_FIELD, color.alpha);
    append_varint_field(
        &mut payload,
        COLOR_SPACE_FIELD,
        match color.rgb_space {
            RgbSpace::Srgb => SRGB_SPACE as u64,
            RgbSpace::DisplayP3 => P3_SPACE as u64,
        },
    );
    Ok(payload)
}

fn rewrite_solid_fill(
    source: &[u8],
    _snapshot: BackgroundSnapshot<'_>,
    color: ColorWrite,
    max_message_bytes: usize,
) -> Result<Vec<u8>, DecodeError> {
    validate_color_write(color)?;
    let replacement = encode_color(color, max_message_bytes)?;
    let mut output = try_vec_bounded(
        source
            .len()
            .saturating_add(replacement.len())
            .saturating_add(32),
        max_message_bytes,
    )?;
    let mut input = source;
    let mut found_color = false;
    while !input.is_empty() {
        let before = input.len();
        let field = next_field(&mut input)?.ok_or_else(DecodeError::projection)?;
        let consumed = before - input.len();
        let raw = &source[source.len() - before..source.len() - before + consumed];
        match field.number {
            FILL_COLOR_FIELD => {
                if found_color {
                    return Err(DecodeError::duplicate_singular("TSD.FillArchive.color"));
                }
                found_color = true;
                let payload = field.bytes()?;
                let rewritten = rewrite_color_payload(payload, color, max_message_bytes)?;
                append_length_field(&mut output, FILL_COLOR_FIELD, &rewritten);
            },
            // A solid replacement cannot retain another selected fill branch,
            // but all future/extension root fields retain exact source bytes.
            FILL_GRADIENT_FIELD | FILL_IMAGE_FIELD => {},
            _ => output.extend_from_slice(raw),
        }
    }
    if !found_color {
        append_length_field(&mut output, FILL_COLOR_FIELD, &replacement);
    }
    Ok(output)
}

/// Replace selected native gradient fields while retaining every unknown
/// `FillArchive` extension field and its original wire framing.
fn rewrite_gradient_fill(
    source: &[u8],
    gradient: GradientWrite<'_>,
    max_message_bytes: usize,
) -> Result<Vec<u8>, DecodeError> {
    let replacement = encode_gradient_payload(gradient, max_message_bytes)?;
    let mut output = try_vec_bounded(
        source
            .len()
            .saturating_add(replacement.len())
            .saturating_add(32),
        max_message_bytes,
    )?;
    let mut input = source;
    let mut found_gradient = false;
    while !input.is_empty() {
        let before = input.len();
        let field = next_field(&mut input)?.ok_or_else(DecodeError::projection)?;
        let consumed = before - input.len();
        let raw = &source[source.len() - before..source.len() - before + consumed];
        match field.number {
            FILL_COLOR_FIELD | FILL_IMAGE_FIELD => {},
            FILL_GRADIENT_FIELD => {
                if found_gradient {
                    return Err(DecodeError::duplicate_singular("TSD.FillArchive.gradient"));
                }
                found_gradient = true;
                let payload = field.bytes()?;
                reject_unknown_selected_gradient_fields(payload)?;
                let rewritten = rewrite_gradient_payload(payload, &replacement, max_message_bytes)?;
                append_length_field(&mut output, FILL_GRADIENT_FIELD, &rewritten);
            },
            _ => output.extend_from_slice(raw),
        }
    }
    if !found_gradient {
        append_length_field(&mut output, FILL_GRADIENT_FIELD, &replacement);
    }
    Ok(output)
}

/// A typed gradient replacement cannot map arbitrary extension fields nested
/// inside a stop, its color, or its angle archive to a new stop/value.  Refuse
/// that rewrite instead of silently discarding producer-authored bytes.
fn reject_unknown_selected_gradient_fields(source: &[u8]) -> Result<(), DecodeError> {
    let mut input = source;
    while let Some(field) = next_field(&mut input)? {
        match field.number {
            GRADIENT_STOP_FIELD => {
                let mut stop_input = field.bytes()?;
                while let Some(stop_field) = next_field(&mut stop_input)? {
                    match stop_field.number {
                        STOP_COLOR_FIELD => {
                            let mut color_input = stop_field.bytes()?;
                            while let Some(color_field) = next_field(&mut color_input)? {
                                if !matches!(
                                    color_field.number,
                                    COLOR_MODEL_FIELD
                                        | COLOR_RED_FIELD..=COLOR_BLUE_FIELD
                                        | COLOR_ALPHA_FIELD
                                        | COLOR_CYAN_FIELD..=COLOR_WHITE_FIELD
                                        | COLOR_SPACE_FIELD
                                ) {
                                    return Err(DecodeError::projection());
                                }
                            }
                        },
                        STOP_FRACTION_FIELD | STOP_INFLECTION_FIELD => {},
                        _ => return Err(DecodeError::projection()),
                    }
                }
            },
            GRADIENT_TYPE_FIELD
            | GRADIENT_OPACITY_FIELD
            | GRADIENT_ADVANCED_FIELD
            | GRADIENT_ANGLE_FIELD => {
                if field.number == GRADIENT_ANGLE_FIELD {
                    let mut angle_input = field.bytes()?;
                    while let Some(angle_field) = next_field(&mut angle_input)? {
                        if angle_field.number != ANGLE_RADIANS_FIELD {
                            return Err(DecodeError::projection());
                        }
                    }
                }
            },
            // A typed replacement has no representation for the transform
            // archive.  Retaining the field would attach it to a new set of
            // stops, while dropping it would discard producer-authored
            // geometry.  Refuse every transform, including transforms that
            // contain only unknown nested extensions.
            GRADIENT_TRANSFORM_FIELD => return Err(DecodeError::projection()),
            _ => {},
        }
    }
    Ok(())
}

/// Replace selected `GradientArchive` fields while retaining unknown nested
/// extension fields with their original non-canonical framing. Repeated stop
/// messages are selected semantic data and are intentionally replaced as one
/// unit; an extension field inside an old stop therefore cannot silently
/// attach to a new stop value.
fn rewrite_gradient_payload(
    source: &[u8],
    replacement: &[u8],
    max_message_bytes: usize,
) -> Result<Vec<u8>, DecodeError> {
    let mut unknown = try_vec_bounded(source.len(), max_message_bytes)?;
    let mut input = source;
    while !input.is_empty() {
        let before = input.len();
        let field = next_field(&mut input)?.ok_or_else(DecodeError::projection)?;
        let consumed = before - input.len();
        let raw = &source[source.len() - before..source.len() - before + consumed];
        if field.number > GRADIENT_TRANSFORM_FIELD {
            unknown.extend_from_slice(raw);
        }
    }
    let mut output = try_vec_bounded(
        replacement
            .len()
            .saturating_add(unknown.len())
            .saturating_add(16),
        max_message_bytes,
    )?;
    output.extend_from_slice(replacement);
    output.extend_from_slice(&unknown);
    Ok(output)
}

/// Patch the known RGB color scalars in place, dropping incompatible CMYK or
/// white components while retaining unknown color extension fields verbatim.
fn rewrite_color_payload(
    source: &[u8],
    color: ColorWrite,
    max_message_bytes: usize,
) -> Result<Vec<u8>, DecodeError> {
    validate_color_write(color)?;
    let mut output = try_vec_bounded(source.len().saturating_add(32), max_message_bytes)?;
    let mut input = source;
    let mut seen = 0u16;
    while !input.is_empty() {
        let before = input.len();
        let field = next_field(&mut input)?.ok_or_else(DecodeError::projection)?;
        let consumed = before - input.len();
        let raw = &source[source.len() - before..source.len() - before + consumed];
        match field.number {
            COLOR_MODEL_FIELD => {
                unique_bit(&mut seen, 0, "TSP.Color.model")?;
                field.int32()?;
                append_varint_field(&mut output, COLOR_MODEL_FIELD, RGB_MODEL as u64);
            },
            COLOR_RED_FIELD => {
                unique_bit(&mut seen, 1, "TSP.Color.r")?;
                field.float()?;
                append_fixed32_field(&mut output, COLOR_RED_FIELD, color.red);
            },
            COLOR_GREEN_FIELD => {
                unique_bit(&mut seen, 2, "TSP.Color.g")?;
                field.float()?;
                append_fixed32_field(&mut output, COLOR_GREEN_FIELD, color.green);
            },
            COLOR_BLUE_FIELD => {
                unique_bit(&mut seen, 3, "TSP.Color.b")?;
                field.float()?;
                append_fixed32_field(&mut output, COLOR_BLUE_FIELD, color.blue);
            },
            COLOR_ALPHA_FIELD => {
                unique_bit(&mut seen, 4, "TSP.Color.a")?;
                field.float()?;
                append_fixed32_field(&mut output, COLOR_ALPHA_FIELD, color.alpha);
            },
            COLOR_SPACE_FIELD => {
                unique_bit(&mut seen, 5, "TSP.Color.rgbspace")?;
                field.int32()?;
                append_varint_field(
                    &mut output,
                    COLOR_SPACE_FIELD,
                    match color.rgb_space {
                        RgbSpace::Srgb => SRGB_SPACE as u64,
                        RgbSpace::DisplayP3 => P3_SPACE as u64,
                    },
                );
            },
            COLOR_CYAN_FIELD | COLOR_MAGENTA_FIELD | COLOR_YELLOW_FIELD | COLOR_BLACK_FIELD
            | COLOR_WHITE_FIELD => {
                // These are known selected fields, not extensions.  They are
                // intentionally removed when a semantic RGB solid is written.
            },
            _ => output.extend_from_slice(raw),
        }
    }
    if seen & (1 << 0) == 0 {
        append_varint_field(&mut output, COLOR_MODEL_FIELD, RGB_MODEL as u64);
    }
    if seen & (1 << 1) == 0 {
        append_fixed32_field(&mut output, COLOR_RED_FIELD, color.red);
    }
    if seen & (1 << 2) == 0 {
        append_fixed32_field(&mut output, COLOR_GREEN_FIELD, color.green);
    }
    if seen & (1 << 3) == 0 {
        append_fixed32_field(&mut output, COLOR_BLUE_FIELD, color.blue);
    }
    if seen & (1 << 4) == 0 {
        append_fixed32_field(&mut output, COLOR_ALPHA_FIELD, color.alpha);
    }
    if seen & (1 << 5) == 0 {
        append_varint_field(
            &mut output,
            COLOR_SPACE_FIELD,
            match color.rgb_space {
                RgbSpace::Srgb => SRGB_SPACE as u64,
                RgbSpace::DisplayP3 => P3_SPACE as u64,
            },
        );
    }
    Ok(output)
}

fn try_vec(capacity: usize) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_error| DecodeError {
            kind: DecodeErrorKind::Allocation {
                requested: capacity,
            },
        })?;
    Ok(output)
}

fn try_vec_bounded(capacity: usize, maximum: usize) -> Result<Vec<u8>, DecodeError> {
    // The exact output size is checked before encoding.  Capping the initial
    // reservation here keeps all handwritten scratch buffers within the
    // configured message ceiling while still allowing a measured candidate
    // smaller than that ceiling to grow normally.
    try_vec(capacity.min(maximum))
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    append_varint(output, u64::from(number) << 3);
    append_varint(output, value);
}

fn append_fixed32_field(output: &mut Vec<u8>, number: u32, value: f32) {
    append_varint(output, (u64::from(number) << 3) | 5);
    output.extend_from_slice(&value.to_le_bytes());
}

fn append_length_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    append_varint(output, (u64::from(number) << 3) | 2);
    append_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options() -> DecodeOptions {
        DecodeOptions::new(1 << 20, 16).with_resource_limits(1 << 16, 1 << 24)
    }

    fn color_payload(
        model: i32,
        red: Option<f32>,
        green: Option<f32>,
        blue: Option<f32>,
        alpha: Option<f32>,
        rgb_space: Option<i32>,
    ) -> Vec<u8> {
        let mut output = Vec::new();
        append_varint_field(&mut output, COLOR_MODEL_FIELD, model as u64);
        if let Some(value) = red {
            append_fixed32_field(&mut output, COLOR_RED_FIELD, value);
        }
        if let Some(value) = green {
            append_fixed32_field(&mut output, COLOR_GREEN_FIELD, value);
        }
        if let Some(value) = blue {
            append_fixed32_field(&mut output, COLOR_BLUE_FIELD, value);
        }
        if let Some(value) = alpha {
            append_fixed32_field(&mut output, COLOR_ALPHA_FIELD, value);
        }
        if let Some(value) = rgb_space {
            append_varint_field(&mut output, COLOR_SPACE_FIELD, value as u64);
        }
        output
    }

    fn solid_fill() -> Vec<u8> {
        let color = color_payload(1, Some(0.1), Some(0.2), Some(0.3), Some(0.4), Some(1));
        let mut output = Vec::new();
        append_length_field(&mut output, FILL_COLOR_FIELD, &color);
        output
    }

    fn gradient_payload() -> Vec<u8> {
        gradient_payload_with(0, 0.0, 1.0, 0.5, 0.5)
    }

    fn gradient_payload_with(
        kind: u64,
        first_fraction: f32,
        second_fraction: f32,
        first_inflection: f32,
        second_inflection: f32,
    ) -> Vec<u8> {
        let mut output = Vec::new();
        append_varint_field(&mut output, GRADIENT_TYPE_FIELD, kind);
        for (red, fraction, inflection) in [
            (0.1, first_fraction, first_inflection),
            (0.9, second_fraction, second_inflection),
        ] {
            let color = color_payload(1, Some(red), Some(0.2), Some(0.3), Some(1.0), Some(1));
            let mut stop = Vec::new();
            append_length_field(&mut stop, STOP_COLOR_FIELD, &color);
            append_fixed32_field(&mut stop, STOP_FRACTION_FIELD, fraction);
            append_fixed32_field(&mut stop, STOP_INFLECTION_FIELD, inflection);
            append_length_field(&mut output, GRADIENT_STOP_FIELD, &stop);
        }
        append_fixed32_field(&mut output, GRADIENT_OPACITY_FIELD, 0.75);
        append_varint_field(&mut output, GRADIENT_ADVANCED_FIELD, 0);
        let mut angle = Vec::new();
        append_fixed32_field(&mut angle, ANGLE_RADIANS_FIELD, 1.25);
        append_length_field(&mut output, GRADIENT_ANGLE_FIELD, &angle);
        output
    }

    fn gradient_fill() -> Vec<u8> {
        let mut output = Vec::new();
        append_length_field(&mut output, FILL_GRADIENT_FIELD, &gradient_payload());
        output
    }

    fn advanced_gradient_fill() -> Vec<u8> {
        let mut payload = Vec::new();
        append_varint_field(&mut payload, GRADIENT_TYPE_FIELD, 0);
        for (red, fraction) in [
            (0.0, 0.0),
            (0.1, 0.1),
            (0.2, 0.2),
            (0.3, 0.3),
            (0.4, 0.4),
            (0.5, 0.5),
            (0.6, 0.6),
            (0.7, 0.7),
            (0.8, 0.8),
            (0.9, 0.9),
            (1.0, 1.0),
        ] {
            let color = color_payload(1, Some(red), Some(0.2), Some(0.3), Some(1.0), Some(1));
            let mut stop = Vec::new();
            append_length_field(&mut stop, STOP_COLOR_FIELD, &color);
            append_fixed32_field(&mut stop, STOP_FRACTION_FIELD, fraction);
            append_fixed32_field(&mut stop, STOP_INFLECTION_FIELD, 0.5);
            append_length_field(&mut payload, GRADIENT_STOP_FIELD, &stop);
        }
        append_fixed32_field(&mut payload, GRADIENT_OPACITY_FIELD, 0.75);
        append_varint_field(&mut payload, GRADIENT_ADVANCED_FIELD, 1);
        let mut angle = Vec::new();
        append_fixed32_field(&mut angle, ANGLE_RADIANS_FIELD, 1.25);
        append_length_field(&mut payload, GRADIENT_ANGLE_FIELD, &angle);

        root_field(FILL_GRADIENT_FIELD, &payload)
    }

    fn gradient_payload_with_transform() -> Vec<u8> {
        let mut payload = gradient_payload();
        let mut transform = Vec::new();
        append_unknown(&mut transform);
        append_length_field(&mut payload, GRADIENT_TRANSFORM_FIELD, &transform);
        payload
    }

    fn gradient_payload_with_nested_unknown_stop() -> Vec<u8> {
        let mut output = Vec::new();
        append_varint_field(&mut output, GRADIENT_TYPE_FIELD, 0);
        for (red, fraction, add_unknown) in [(0.1, 0.0, true), (0.9, 1.0, false)] {
            let color = color_payload(1, Some(red), Some(0.2), Some(0.3), Some(1.0), Some(1));
            let mut stop = Vec::new();
            append_length_field(&mut stop, STOP_COLOR_FIELD, &color);
            append_fixed32_field(&mut stop, STOP_FRACTION_FIELD, fraction);
            append_fixed32_field(&mut stop, STOP_INFLECTION_FIELD, 0.5);
            if add_unknown {
                append_unknown(&mut stop);
            }
            append_length_field(&mut output, GRADIENT_STOP_FIELD, &stop);
        }
        append_fixed32_field(&mut output, GRADIENT_OPACITY_FIELD, 0.75);
        append_varint_field(&mut output, GRADIENT_ADVANCED_FIELD, 0);
        let mut angle = Vec::new();
        append_fixed32_field(&mut angle, ANGLE_RADIANS_FIELD, 1.25);
        append_length_field(&mut output, GRADIENT_ANGLE_FIELD, &angle);
        output
    }

    fn root_field(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        append_length_field(&mut output, number, payload);
        output
    }

    fn overlong_unknown_varint_field(number: u32, value: u8) -> Vec<u8> {
        let tag = u64::from(number) << 3;
        let mut output = Vec::new();
        append_varint(&mut output, tag);
        let last = output.len() - 1;
        output[last] |= 0x80;
        output.push(0);
        output.push(value);
        output
    }

    fn overlong_unknown_length_field(number: u32, payload: &[u8]) -> Vec<u8> {
        let tag = (u64::from(number) << 3) | 2;
        let mut output = Vec::new();
        append_varint(&mut output, tag);
        let last = output.len() - 1;
        output[last] |= 0x80;
        output.push(0);
        output.extend_from_slice(&[0x81, 0x00]);
        output.extend_from_slice(payload);
        output
    }

    fn append_unknown(output: &mut Vec<u8>) -> Vec<u8> {
        let unknown = overlong_unknown_varint_field(100, 7);
        output.extend_from_slice(&unknown);
        unknown
    }

    fn typed_solid() -> BackgroundWrite<'static> {
        BackgroundWrite::Solid(ColorWrite::new(0.1, 0.2, 0.3, 0.4, RgbSpace::Srgb))
    }

    fn typed_gradient<'source>(stops: &'source [GradientStopWrite]) -> BackgroundWrite<'source> {
        BackgroundWrite::Gradient(GradientWrite {
            kind: GradientKind::Linear,
            stops,
            opacity: 0.75,
            advanced: false,
            angle_radians: 1.25,
        })
    }

    #[test]
    fn empty_solid_image_and_opaque_payloads_project_lazily() {
        assert!(matches!(
            decode_slide_background(&[], options()).unwrap(),
            BackgroundSnapshot::None { .. }
        ));
        assert!(matches!(
            decode_slide_background(&solid_fill(), options()).unwrap(),
            BackgroundSnapshot::Solid { .. }
        ));
        assert!(matches!(
            decode_slide_background(&root_field(FILL_IMAGE_FIELD, &[]), options()).unwrap(),
            BackgroundSnapshot::Image { .. }
        ));
        let opaque = overlong_unknown_varint_field(100, 7);
        assert!(matches!(
            decode_slide_background(&opaque, options()).unwrap(),
            BackgroundSnapshot::Opaque { .. }
        ));
    }

    #[test]
    fn required_duplicate_wire_and_finite_color_fields_are_strict() {
        let mut missing_model = Vec::new();
        append_fixed32_field(&mut missing_model, COLOR_RED_FIELD, 0.1);
        append_fixed32_field(&mut missing_model, COLOR_GREEN_FIELD, 0.2);
        append_fixed32_field(&mut missing_model, COLOR_BLUE_FIELD, 0.3);
        append_varint_field(&mut missing_model, COLOR_SPACE_FIELD, 1);
        let error =
            decode_slide_background(&root_field(FILL_COLOR_FIELD, &missing_model), options())
                .unwrap_err();
        assert_eq!(error.missing_required_field(), Some("TSP.Color.model"));

        let incomplete_rgb = encode_varint_field_for_test(COLOR_MODEL_FIELD, RGB_MODEL as u64);
        let error =
            decode_slide_background(&root_field(FILL_COLOR_FIELD, &incomplete_rgb), options())
                .unwrap_err();
        assert_eq!(error.missing_required_field(), Some("TSP.Color.r"));

        let mut duplicate = color_payload(1, Some(0.1), Some(0.2), Some(0.3), None, Some(1));
        append_varint_field(&mut duplicate, COLOR_MODEL_FIELD, 1);
        let error = decode_slide_background(&root_field(FILL_COLOR_FIELD, &duplicate), options())
            .unwrap_err();
        assert_eq!(error.duplicate_singular_field(), Some("TSP.Color.model"));

        let mut wrong_wire = Vec::new();
        append_fixed32_field(&mut wrong_wire, COLOR_MODEL_FIELD, 1.0);
        let error = decode_slide_background(&root_field(FILL_COLOR_FIELD, &wrong_wire), options())
            .unwrap_err();
        assert!(matches!(error.kind, DecodeErrorKind::Wire(_)));

        let nan = color_payload(1, Some(f32::NAN), Some(0.2), Some(0.3), None, Some(1));
        let error =
            decode_slide_background(&root_field(FILL_COLOR_FIELD, &nan), options()).unwrap_err();
        assert_eq!(
            error.noncanonical_reason(),
            Some("float scalar is not finite")
        );

        let out_of_range = color_payload(1, Some(1.01), Some(0.2), Some(0.3), None, Some(1));
        let error =
            decode_slide_background(&root_field(FILL_COLOR_FIELD, &out_of_range), options())
                .unwrap_err();
        assert_eq!(
            error.noncanonical_reason(),
            Some("color component outside [0,1]")
        );
    }

    #[test]
    fn gradient_projection_retains_exact_stop_presence_and_optional_incomplete_is_opaque() {
        let source = gradient_fill();
        let (snapshot, report) = decode_slide_background_with_report(&source, options()).unwrap();
        let BackgroundSnapshot::Gradient { gradient, .. } = snapshot else {
            panic!("expected supported gradient");
        };
        assert_eq!(gradient.gradient_type, Some(0));
        assert_eq!(gradient.opacity.map(f32::to_bits), Some(0.75f32.to_bits()));
        let stops = gradient
            .stops(options())
            .collect::<Result<Vec<_>, _>>()
            .unwrap();
        assert_eq!(stops.len(), 2);
        assert_eq!(stops[0].fraction, Some(0.0));
        assert_eq!(stops[1].fraction, Some(1.0));
        assert!(report.fields > 8);

        let incomplete = root_field(FILL_GRADIENT_FIELD, &encode_varint_field_for_test(1, 0));
        assert!(matches!(
            decode_slide_background(&incomplete, options()).unwrap(),
            BackgroundSnapshot::Opaque { .. }
        ));

        let mut mixed = solid_fill();
        mixed.extend_from_slice(&gradient_fill());
        assert!(matches!(
            decode_slide_background(&mixed, options()).unwrap(),
            BackgroundSnapshot::Opaque { .. }
        ));
    }

    #[test]
    fn for_source_allows_a_large_valid_gradient_within_its_work_ceiling() {
        let source = advanced_gradient_fill();
        let options = DecodeOptions::for_source(&source);
        let (snapshot, report) = decode_slide_background_with_report(&source, options).unwrap();
        assert!(matches!(snapshot, BackgroundSnapshot::Gradient { .. }));
        assert!(report.work_bytes <= source.len().saturating_mul(16));
    }

    #[test]
    fn unsupported_simple_gradient_shapes_are_opaque_after_strict_scan() {
        for payload in [
            gradient_payload_with(0, 0.75, 0.25, 0.5, 0.5),
            gradient_payload_with(1, 0.0, 1.0, 0.5, 0.5),
            gradient_payload_with(0, 0.0, 1.0, 0.4, 0.5),
        ] {
            let source = root_field(FILL_GRADIENT_FIELD, &payload);
            assert!(matches!(
                decode_slide_background(&source, options()).unwrap(),
                BackgroundSnapshot::Opaque { .. }
            ));
        }
    }

    #[test]
    fn image_nested_required_references_are_checked() {
        let missing_reference = root_field(
            FILL_IMAGE_FIELD,
            &root_field(IMAGE_DATABASE_DATA_FIELD, &[]),
        );
        let error = decode_slide_background(&missing_reference, options()).unwrap_err();
        assert_eq!(
            error.missing_required_field(),
            Some("TSP.Reference.identifier")
        );

        let reference = encode_varint_field_for_test(RESOURCE_IDENTIFIER_FIELD, 12);
        let mut duplicate = reference.clone();
        duplicate.extend_from_slice(&reference);
        let error = decode_slide_background(
            &root_field(
                FILL_IMAGE_FIELD,
                &root_field(IMAGE_DATABASE_DATA_FIELD, &duplicate),
            ),
            options(),
        )
        .unwrap_err();
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSP.Reference.identifier")
        );

        let wrong_wire = root_field(
            IMAGE_DATABASE_DATA_FIELD,
            &encode_fixed32_field_for_test(1, 12),
        );
        let error = decode_slide_background(&root_field(FILL_IMAGE_FIELD, &wrong_wire), options())
            .unwrap_err();
        assert!(matches!(error.kind, DecodeErrorKind::Wire(_)));
    }

    #[test]
    fn selected_gradient_fields_remain_strict_even_when_semantics_are_opaque() {
        let mut duplicate_type = encode_varint_field_for_test(GRADIENT_TYPE_FIELD, 0);
        duplicate_type.extend_from_slice(&encode_varint_field_for_test(GRADIENT_TYPE_FIELD, 1));
        let error =
            decode_slide_background(&root_field(FILL_GRADIENT_FIELD, &duplicate_type), options())
                .unwrap_err();
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSD.GradientArchive.type")
        );

        let wrong_wire = encode_fixed32_field_for_test(GRADIENT_ADVANCED_FIELD, 0);
        let error =
            decode_slide_background(&root_field(FILL_GRADIENT_FIELD, &wrong_wire), options())
                .unwrap_err();
        assert!(matches!(error.kind, DecodeErrorKind::Wire(_)));

        let mut nonfinite = Vec::new();
        append_fixed32_field(&mut nonfinite, GRADIENT_OPACITY_FIELD, f32::INFINITY);
        let error =
            decode_slide_background(&root_field(FILL_GRADIENT_FIELD, &nonfinite), options())
                .unwrap_err();
        assert_eq!(
            error.noncanonical_reason(),
            Some("float scalar is not finite")
        );
    }

    #[test]
    fn unknown_noncanonical_root_and_nested_fields_are_accepted_verbatim() {
        let mut color = color_payload(1, Some(0.1), Some(0.2), Some(0.3), None, Some(1));
        let nested_unknown = overlong_unknown_length_field(101, &[0x42]);
        color.extend_from_slice(&nested_unknown);
        let mut source = root_field(FILL_COLOR_FIELD, &color);
        let root_unknown = overlong_unknown_varint_field(100, 7);
        source.extend_from_slice(&root_unknown);
        let (snapshot, _) = decode_slide_background_with_report(&source, options()).unwrap();
        assert!(matches!(snapshot, BackgroundSnapshot::Solid { .. }));
        assert!(
            source
                .windows(root_unknown.len())
                .any(|window| window == root_unknown)
        );
        assert!(
            source
                .windows(nested_unknown.len())
                .any(|window| window == nested_unknown)
        );
    }

    #[test]
    fn groups_are_rejected_at_root_and_nested_boundaries() {
        let root_group = [0x23, 0x24];
        assert!(decode_slide_background(&root_group, options()).is_err());
        let nested_group = [0x23, 0x24];
        let color = color_payload(1, Some(0.1), Some(0.2), Some(0.3), None, Some(1));
        let mut color_with_group = color;
        color_with_group.extend_from_slice(&nested_group);
        assert!(
            decode_slide_background(&root_field(FILL_COLOR_FIELD, &color_with_group), options())
                .is_err()
        );
    }

    #[test]
    fn exact_fields_work_bytes_and_depth_boundaries_are_enforced() {
        let mut source = solid_fill();
        append_unknown(&mut source);
        append_unknown(&mut source);
        let exact_fields = DecodeOptions::new(source.len(), 16).with_resource_limits(32, 1 << 24);
        let (_, report) = decode_slide_background_with_report(&source, exact_fields).unwrap();
        let exact = DecodeOptions::new(source.len(), 16)
            .with_resource_limits(report.fields, report.work_bytes);
        assert!(decode_slide_background(&source, exact).is_ok());
        let field_limited = exact.with_resource_limits(report.fields - 1, report.work_bytes);
        let error = decode_slide_background(&source, field_limited).unwrap_err();
        assert_eq!(
            error.field_limit_values(),
            Some((report.fields, report.fields - 1))
        );
        let work_limited = exact.with_resource_limits(report.fields, report.work_bytes - 1);
        let error = decode_slide_background(&source, work_limited).unwrap_err();
        assert_eq!(
            error.work_limit_values(),
            Some((report.work_bytes, report.work_bytes - 1))
        );
        let bytes_limited =
            DecodeOptions::new(source.len() - 1, 16).with_resource_limits(1 << 16, 1 << 24);
        assert!(matches!(
            decode_slide_background(&source, bytes_limited).unwrap_err().wire_resource_limit(),
            Some(WireResourceLimit::Bytes { observed, maximum })
                if observed == source.len() && maximum == source.len() - 1
        ));

        let gradient = gradient_fill();
        let exact_depth = DecodeOptions::new(1 << 20, 3).with_resource_limits(1 << 16, 1 << 24);
        assert!(decode_slide_background(&gradient, exact_depth).is_ok());
        let one_over = DecodeOptions::new(1 << 20, 2).with_resource_limits(1 << 16, 1 << 24);
        assert!(matches!(
            decode_slide_background(&gradient, one_over)
                .unwrap_err()
                .wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 3,
                maximum: 2
            })
        ));
    }

    #[test]
    fn solid_rewrite_preserves_unknown_root_and_color_framing() {
        let mut color = color_payload(1, Some(0.1), Some(0.2), Some(0.3), Some(0.4), Some(1));
        let color_unknown = overlong_unknown_varint_field(100, 7);
        color.extend_from_slice(&color_unknown);
        let mut source = root_field(FILL_COLOR_FIELD, &color);
        let root_unknown = append_unknown(&mut source);
        let (candidate, _) =
            rewrite_slide_background_with_report(&source, typed_solid(), options()).unwrap();
        assert!(
            candidate
                .windows(root_unknown.len())
                .any(|window| window == root_unknown)
        );
        assert!(
            candidate
                .windows(color_unknown.len())
                .any(|window| window == color_unknown)
        );
        let (noop, report) =
            rewrite_slide_background_with_report(&solid_fill(), typed_solid(), options()).unwrap();
        assert_eq!(noop, solid_fill());
        assert!(!report.changed);
        assert_eq!(report.allocations, 3);
    }

    #[test]
    fn typed_gradient_rewrite_preserves_unknown_root_and_reads_every_stop() {
        let stops = [
            GradientStopWrite {
                color: ColorWrite::new(0.1, 0.2, 0.3, 1.0, RgbSpace::Srgb),
                fraction: 0.0,
                inflection: 0.5,
            },
            GradientStopWrite {
                color: ColorWrite::new(0.9, 0.2, 0.3, 1.0, RgbSpace::Srgb),
                fraction: 0.75,
                inflection: 0.5,
            },
        ];
        let mut source = gradient_fill();
        let unknown = append_unknown(&mut source);
        let (candidate, report) =
            rewrite_slide_background_with_report(&source, typed_gradient(&stops), options())
                .unwrap();
        assert!(report.changed);
        assert_eq!(report.allocations, 9);
        assert!(
            candidate
                .windows(unknown.len())
                .any(|window| window == unknown)
        );
        let (snapshot, _) = decode_slide_background_with_report(&candidate, options()).unwrap();
        let BackgroundSnapshot::Gradient { gradient, .. } = snapshot else {
            panic!("expected gradient readback");
        };
        let actual = gradient
            .stops(options())
            .collect::<Result<Vec<_>, _>>()
            .unwrap();
        assert_eq!(actual.len(), stops.len());
        for (actual, expected) in actual.into_iter().zip(stops) {
            assert_eq!(
                actual.fraction.map(f32::to_bits),
                Some(expected.fraction.to_bits())
            );
            assert_eq!(
                actual.inflection.map(f32::to_bits),
                Some(expected.inflection.to_bits())
            );
            assert!(
                actual
                    .color
                    .is_some_and(|actual| same_color_value(actual, expected.color))
            );
        }
    }

    #[test]
    fn typed_gradient_rewrite_refuses_unmappable_nested_extensions() {
        let nested_gradient = gradient_payload_with_nested_unknown_stop();
        let source = root_field(FILL_GRADIENT_FIELD, &nested_gradient);
        let stops = [
            GradientStopWrite {
                color: ColorWrite::new(0.1, 0.2, 0.3, 1.0, RgbSpace::Srgb),
                fraction: 0.0,
                inflection: 0.5,
            },
            GradientStopWrite {
                color: ColorWrite::new(0.9, 0.2, 0.3, 1.0, RgbSpace::Srgb),
                fraction: 0.75,
                inflection: 0.5,
            },
        ];
        let error =
            rewrite_slide_background_with_report(&source, typed_gradient(&stops), options())
                .unwrap_err();
        assert_eq!(error.allocation_requested(), None);
    }

    #[test]
    fn typed_gradient_rewrite_refuses_transforms_and_nested_extensions() {
        let source = root_field(FILL_GRADIENT_FIELD, &gradient_payload_with_transform());
        let stops = [
            GradientStopWrite {
                color: ColorWrite::new(0.1, 0.2, 0.3, 1.0, RgbSpace::Srgb),
                fraction: 0.0,
                inflection: 0.5,
            },
            GradientStopWrite {
                color: ColorWrite::new(0.9, 0.2, 0.3, 1.0, RgbSpace::Srgb),
                fraction: 0.75,
                inflection: 0.5,
            },
        ];
        let error =
            rewrite_slide_background_with_report(&source, typed_gradient(&stops), options())
                .unwrap_err();
        assert_eq!(error.allocation_requested(), None);
    }

    #[test]
    fn rewrite_work_limit_is_aggregate_over_source_encode_and_readback() {
        let source = solid_fill();
        let (_, generous) = rewrite_slide_background_with_report(
            &source,
            typed_solid(),
            DecodeOptions::new(1 << 20, 16).with_resource_limits(1 << 16, 1 << 24),
        )
        .unwrap();
        let aggregate_limit = generous.work_bytes - 1;
        let options =
            DecodeOptions::new(1 << 20, 16).with_resource_limits(1 << 16, aggregate_limit);
        let error =
            rewrite_slide_background_with_report(&source, typed_solid(), options).unwrap_err();
        assert!(error.work_limit_values().is_some());
    }

    #[test]
    fn oversized_option_reports_option_limit_not_source_length() {
        let options = DecodeOptions::new(usize::MAX, 16);
        let error = decode_slide_background(&[], options).unwrap_err();
        assert_eq!(
            error.wire_resource_limit(),
            Some(WireResourceLimit::Bytes {
                observed: usize::MAX,
                maximum: usize::try_from(buffa::MAX_MESSAGE_BYTES).unwrap(),
            })
        );
    }

    #[test]
    fn rewrite_buffer_reservation_reports_allocation_mapping() {
        let error = try_vec(usize::MAX).unwrap_err();
        assert_eq!(error.allocation_requested(), Some(usize::MAX));
    }

    #[test]
    fn bounded_scratch_reservation_respects_message_ceiling() {
        let output = try_vec_bounded(usize::MAX, 17).unwrap();
        assert!(output.capacity() <= 17);
    }

    #[test]
    fn clear_and_raw_rewrites_count_handwritten_buffers() {
        let options = options();
        let (_, clear_report) =
            rewrite_slide_background_with_report(&solid_fill(), BackgroundWrite::Clear, options)
                .unwrap();
        assert_eq!(clear_report.allocations, 0);
        let raw = solid_fill();
        let (_, raw_report) =
            rewrite_slide_background_with_report(&[], BackgroundWrite::Raw(&raw), options).unwrap();
        assert_eq!(raw_report.allocations, 1);
    }

    fn encode_varint_field_for_test(number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        append_varint_field(&mut output, number, value);
        output
    }

    fn encode_fixed32_field_for_test(number: u32, value: u32) -> Vec<u8> {
        let mut output = Vec::new();
        append_varint(&mut output, (u64::from(number) << 3) | 5);
        output.extend_from_slice(&value.to_le_bytes());
        output
    }
}
