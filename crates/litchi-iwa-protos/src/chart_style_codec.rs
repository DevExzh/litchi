//! Bounded wire codec for the generated property extensions of the TSCH
//! chart style archives, and the style-reference projection of a chart.
//!
//! Every chart owns eight style objects: the chart, legend, axis and series
//! style and non-style archives (`TSCH.Chart*StyleArchive`,
//! `TSCH.Chart*NonStyleArchive`). Their inspector properties are stored in a
//! generated extension (field 10000) whose message lives in
//! `TSCHArchives.GEN.proto`. This module reads and rewrites individual
//! properties of that extension directly on the wire:
//!
//! - [`decode`](crate::chart_style_codec::decode) validates the outer framing and the extension once and
//!   returns borrowed [`Properties`](crate::chart_style_codec::Properties);
//! - [`rewrite`](crate::chart_style_codec::rewrite) applies a batch of [`Write`](crate::chart_style_codec::Write) values and returns the candidate
//!   payload, borrowing the source when nothing changed.
//!
//! Byte semantics are those of the migration host's per-property patches
//! (record 0375): a field present in the source is replaced in place (its key
//! bytes, and its length prefix when the length is unchanged, are kept); a
//! cleared field is removed; a missing field is appended to the end of the
//! extension; a missing extension is created, with its fields in field-number
//! order, only when the batch sets at least one value. Every unrelated byte,
//! including unknown fields, stays where it was.
//!
//! [`references`](crate::chart_style_codec::references) projects the style and paragraph-style references of one
//! `TSCH.ChartDrawableArchive` payload without decoding its data grid, so the
//! shared chart engine can locate the objects it edits.
//!
//! The codec never resolves objects, allocates identifiers or touches an
//! archive; those belong to the format owners. Native field numbers are
//! exposed only through [`field`](crate::chart_style_codec::field), for the owners' private adapters.

use std::borrow::Cow;
use std::fmt;

/// Field number of the generated property extension in every style archive.
pub const EXTENSION_FIELD: u32 = 10_000;

const WIRE_VARINT: u8 = 0;
const WIRE_FIXED64: u8 = 1;
const WIRE_LENGTH_DELIMITED: u8 = 2;
const WIRE_FIXED32: u8 = 5;
const MAX_FIELD_NUMBER: u64 = (1 << 29) - 1;
const HARD_MAX_BYTES: usize = 256 * 1024 * 1024;
const HARD_MAX_FIELDS: usize = 16 * 1024 * 1024;

/// Native field numbers of the generated extensions and of the chart
/// archive's style references.
pub mod field {
    /// `TSCH.Generated.ChartNonStyleArchive`.
    pub mod chart_non_style {
        /// `tschchartinfodefault3dintersetdepthgap` (float ratio).
        pub const SERIES_GAP_3D: u32 = 3;
        /// `tschchartinfodefault3drotation` (vector, degrees in X and Y).
        pub const ROTATION_3D: u32 = 4;
        /// `tschchartinfoarea3dscale` (vector).
        pub const AREA_SCALE_3D: u32 = 5;
        /// `tschchartinfobar3dscale` (vector).
        pub const BAR_SCALE_3D: u32 = 6;
        /// `tschchartinfocolumn3dscale` (vector).
        pub const COLUMN_SCALE_3D: u32 = 7;
        /// `tschchartinfoline3dscale` (vector).
        pub const LINE_SCALE_3D: u32 = 9;
        /// `tschchartinfopie3dscale` (vector).
        pub const PIE_SCALE_3D: u32 = 10;
        /// `tschchartinfodefaultpiestartangle` (float degrees).
        pub const PIE_START_ANGLE: u32 = 19;
        /// `tschchartinfodefaultinnerradius` (float fraction).
        pub const DONUT_INNER_RADIUS: u32 = 27;
        /// `tschchartinfodefaultradarstartangle` (float degrees).
        pub const RADAR_START_ANGLE: u32 = 29;
    }

    /// `TSCH.Generated.ChartStyleArchive`.
    pub mod chart_style {
        /// `tschchartinfodefaultroundedcornerradius` (float fraction).
        pub const ROUNDED_CORNER_RADIUS: u32 = 122;
        /// `tschchartinfodefaultroundedcornerouterendonly` (bool).
        pub const ROUNDED_CORNER_OUTER_END_ONLY: u32 = 123;
    }

    /// `TSCH.Generated.ChartSeriesNonStyleArchive`.
    pub mod series_non_style {
        /// `tschchartseriespie2_3labelexplosion` (legacy float).
        pub const PIE_LEGACY_LABEL_EXPLOSION: u32 = 16;
        /// `tschchartseriespiewedgeexplosion` (float fraction).
        pub const PIE_WEDGE_EXPLOSION: u32 = 63;
        /// `tschchartseriespielabelexplosion` (float percent).
        pub const PIE_LABEL_EXPLOSION: u32 = 147;
    }

    /// `TSCH.Chart3DVectorArchive` components (four required floats).
    pub mod vector {
        /// `x`.
        pub const X: u32 = 1;
        /// `y`.
        pub const Y: u32 = 2;
        /// `z`.
        pub const Z: u32 = 3;
        /// `w`.
        pub const W: u32 = 4;
    }
}

/// The eight TSCH style archives that carry a generated extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Archive {
    /// `TSCH.ChartStyleArchive` (5022).
    ChartStyle,
    /// `TSCH.ChartNonStyleArchive` (5023).
    ChartNonStyle,
    /// `TSCH.LegendStyleArchive` (5024).
    LegendStyle,
    /// `TSCH.LegendNonStyleArchive` (5025).
    LegendNonStyle,
    /// `TSCH.ChartAxisStyleArchive` (5026).
    AxisStyle,
    /// `TSCH.ChartAxisNonStyleArchive` (5027).
    AxisNonStyle,
    /// `TSCH.ChartSeriesStyleArchive` (5028).
    SeriesStyle,
    /// `TSCH.ChartSeriesNonStyleArchive` (5029).
    SeriesNonStyle,
}

impl Archive {
    /// Every archive, in message-type order.
    pub const ALL: [Self; 8] = [
        Self::ChartStyle,
        Self::ChartNonStyle,
        Self::LegendStyle,
        Self::LegendNonStyle,
        Self::AxisStyle,
        Self::AxisNonStyle,
        Self::SeriesStyle,
        Self::SeriesNonStyle,
    ];

    /// The IWA message type of this archive.
    #[must_use]
    pub const fn message_type(self) -> u32 {
        match self {
            Self::ChartStyle => 5_022,
            Self::ChartNonStyle => 5_023,
            Self::LegendStyle => 5_024,
            Self::LegendNonStyle => 5_025,
            Self::AxisStyle => 5_026,
            Self::AxisNonStyle => 5_027,
            Self::SeriesStyle => 5_028,
            Self::SeriesNonStyle => 5_029,
        }
    }

    /// The archive stored under an IWA message type.
    #[must_use]
    pub const fn from_message_type(message_type: u32) -> Option<Self> {
        match message_type {
            5_022 => Some(Self::ChartStyle),
            5_023 => Some(Self::ChartNonStyle),
            5_024 => Some(Self::LegendStyle),
            5_025 => Some(Self::LegendNonStyle),
            5_026 => Some(Self::AxisStyle),
            5_027 => Some(Self::AxisNonStyle),
            5_028 => Some(Self::SeriesStyle),
            5_029 => Some(Self::SeriesNonStyle),
            _ => None,
        }
    }
}

/// A `TSCH.Chart3DVectorArchive`: four floats, missing components read as 0.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Vector {
    /// `x` (field 1).
    pub x: f32,
    /// `y` (field 2).
    pub y: f32,
    /// `z` (field 3).
    pub z: f32,
    /// `w` (field 4).
    pub w: f32,
}

/// One property write of a [`rewrite`] batch.
///
/// The plain variants set (`Some`) or clear (`None`) one field. The
/// conditional variants carry the presence-dependent rules of the native
/// inspector properties, evaluated against the extension as the batch has
/// left it so far (an absent extension has every field absent).
#[derive(Debug, Clone, Copy, PartialEq)]
#[non_exhaustive]
pub enum Write<'a> {
    /// A `float` (fixed32) property.
    Float {
        /// Extension field number.
        field: u32,
        /// New value, or `None` to remove the field.
        value: Option<f32>,
    },
    /// A `bool` (varint) property.
    Bool {
        /// Extension field number.
        field: u32,
        /// New value, or `None` to remove the field.
        value: Option<bool>,
    },
    /// An `int32` (varint) property.
    Int32 {
        /// Extension field number.
        field: u32,
        /// New value, or `None` to remove the field.
        value: Option<i32>,
    },
    /// Set a `float` property, except when it is absent and `value ==
    /// default` (then nothing is written).
    FloatUnlessAbsentDefault {
        /// Extension field number.
        field: u32,
        /// New value.
        value: f32,
        /// The value an absent field reads as.
        default: f32,
    },
    /// Set a `bool` property, except when it is absent and `value ==
    /// default` (then nothing is written).
    BoolUnlessAbsentDefault {
        /// Extension field number.
        field: u32,
        /// New value.
        value: bool,
        /// The value an absent field reads as.
        default: bool,
    },
    /// A whole `TSCH.Chart3DVectorArchive` property, written with all four
    /// components.
    Vector {
        /// Extension field number.
        field: u32,
        /// New value, or `None` to remove the field.
        value: Option<Vector>,
    },
    /// When the vector is present, replace the given components and keep the
    /// others' bytes; when it is absent, write `absent` whole (or nothing).
    VectorComponents {
        /// Extension field number of the vector.
        field: u32,
        /// New `[x, y, z, w]` components; `None` keeps a component.
        components: [Option<f32>; 4],
        /// The whole vector to write when the field is absent.
        absent: Option<Vector>,
    },
    /// A message-typed property (a stroke or fill payload), replaced as a
    /// whole; `None` removes the field.
    Message {
        /// Extension field number.
        field: u32,
        /// The encoded message, or `None` to remove the field.
        value: Option<&'a [u8]>,
    },
    /// A message-typed property merged into a present field: the fields
    /// `schema` knows are replaced from `value` (singular messages
    /// recursively), unknown fields stay at every level; an absent field is
    /// set to `value`.
    MergeMessage {
        /// Extension field number.
        field: u32,
        /// The encoded replacement message.
        value: &'a [u8],
        /// The schema of the message's known fields.
        schema: &'static [merge::KnownField],
    },
    /// A depth stored as the `z` component of a scale vector: when the vector
    /// is present, `z = ratio * denominator` where the denominator is the
    /// stored `denominator` component ([`field::vector::X`] or
    /// [`field::vector::Y`]), which must be finite and positive; when it is
    /// absent, write `absent` whole (or nothing).
    VectorDepth {
        /// Extension field number of the scale vector.
        field: u32,
        /// The depth ratio.
        ratio: f32,
        /// The component the ratio is relative to.
        denominator: u32,
        /// The whole vector to write when the field is absent.
        absent: Option<Vector>,
    },
}

impl Write<'_> {
    const fn field(self) -> u32 {
        match self {
            Self::Float { field, .. }
            | Self::Bool { field, .. }
            | Self::Int32 { field, .. }
            | Self::FloatUnlessAbsentDefault { field, .. }
            | Self::BoolUnlessAbsentDefault { field, .. }
            | Self::Vector { field, .. }
            | Self::VectorComponents { field, .. }
            | Self::VectorDepth { field, .. }
            | Self::Message { field, .. }
            | Self::MergeMessage { field, .. } => field,
        }
    }

    const fn wire_type(self) -> u8 {
        match self {
            Self::Float { .. } | Self::FloatUnlessAbsentDefault { .. } => WIRE_FIXED32,
            Self::Bool { .. } | Self::Int32 { .. } | Self::BoolUnlessAbsentDefault { .. } => {
                WIRE_VARINT
            },
            Self::Vector { .. }
            | Self::VectorComponents { .. }
            | Self::VectorDepth { .. }
            | Self::Message { .. }
            | Self::MergeMessage { .. } => WIRE_LENGTH_DELIMITED,
        }
    }
}

/// What [`rewrite`] does with a generated extension its writes leave empty.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum EmptyExtension {
    /// Keep the empty extension field (the chart-level properties).
    #[default]
    Keep,
    /// Remove the extension field (the per-series label distances).
    Remove,
}

/// Finite resources a decode or rewrite may use.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_bytes: usize,
    max_fields: usize,
}

impl Limits {
    /// The hard ceilings: 256 MiB payloads, 16 Mi field visits.
    pub const DEFAULT: Self = Self {
        max_bytes: HARD_MAX_BYTES,
        max_fields: HARD_MAX_FIELDS,
    };

    /// Replace the payload-byte ceiling (source and candidate).
    ///
    /// # Errors
    ///
    /// [`Error::InvalidLimit`] for 0 or a value above the hard ceiling.
    pub const fn with_max_bytes(mut self, maximum: usize) -> Result<Self, Error> {
        if maximum == 0 || maximum > HARD_MAX_BYTES {
            return Err(Error::InvalidLimit {
                kind: LimitKind::Bytes,
            });
        }
        self.max_bytes = maximum;
        Ok(self)
    }

    /// Replace the field-visit ceiling.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidLimit`] for 0 or a value above the hard ceiling.
    pub const fn with_max_fields(mut self, maximum: usize) -> Result<Self, Error> {
        if maximum == 0 || maximum > HARD_MAX_FIELDS {
            return Err(Error::InvalidLimit {
                kind: LimitKind::Fields,
            });
        }
        self.max_fields = maximum;
        Ok(self)
    }

    /// The payload-byte ceiling.
    #[must_use]
    pub const fn max_bytes(self) -> usize {
        self.max_bytes
    }

    /// The field-visit ceiling.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// The resource a [`Error::Limit`] names.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Source or candidate payload bytes.
    Bytes,
    /// Wire fields visited.
    Fields,
}

/// A codec failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The payload is not valid protobuf wire data.
    Malformed {
        /// What is wrong.
        reason: &'static str,
    },
    /// A singular field occurs more than once.
    DuplicateField {
        /// Field number.
        field: u32,
    },
    /// A field has the wrong wire type for its schema.
    WireType {
        /// Field number.
        field: u32,
        /// Wire type the schema requires.
        expected: u8,
        /// Wire type found.
        found: u8,
    },
    /// A [`Write::VectorDepth`] found a stored scale vector whose
    /// denominator component is not finite and positive.
    InvalidScale {
        /// Extension field number of the vector.
        field: u32,
    },
    /// A write names field 0, the extension field itself, or a number above
    /// the protobuf maximum.
    InvalidField {
        /// Field number.
        field: u32,
    },
    /// A finite resource was exceeded.
    Limit {
        /// The resource.
        kind: LimitKind,
        /// Amount required.
        observed: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// A limit builder received 0 or a value above the hard ceiling.
    InvalidLimit {
        /// The resource.
        kind: LimitKind,
    },
    /// An output buffer could not be allocated.
    Allocation {
        /// Bytes requested.
        bytes: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match *self {
            Self::Malformed { reason } => {
                write!(formatter, "malformed chart style payload: {reason}")
            },
            Self::DuplicateField { field } => {
                write!(formatter, "chart style field {field} occurs more than once")
            },
            Self::WireType {
                field,
                expected,
                found,
            } => write!(
                formatter,
                "chart style field {field} has wire type {found}; the schema requires {expected}"
            ),
            Self::InvalidScale { field } => write!(
                formatter,
                "chart style scale vector {field} has no finite positive depth denominator"
            ),
            Self::InvalidField { field } => {
                write!(
                    formatter,
                    "chart style field number {field} cannot be written"
                )
            },
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "chart style codec {kind:?} limit exceeded: {observed} > {maximum}"
            ),
            Self::InvalidLimit { kind } => {
                write!(
                    formatter,
                    "chart style codec {kind:?} limit is out of range"
                )
            },
            Self::Allocation { bytes } => {
                write!(
                    formatter,
                    "chart style codec could not allocate {bytes} bytes"
                )
            },
        }
    }
}

impl std::error::Error for Error {}

/// One wire field located in a payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Field {
    pub(crate) number: u32,
    pub(crate) wire_type: u8,
    pub(crate) start: usize,
    pub(crate) key_end: usize,
    pub(crate) payload_start: usize,
    pub(crate) end: usize,
    /// The decoded varint, for wire type 0.
    pub(crate) varint: u64,
}

impl Field {
    pub(crate) fn payload<'a>(&self, data: &'a [u8]) -> &'a [u8] {
        &data[self.payload_start..self.end]
    }
}

/// Field visits charged against [`Limits::max_fields`].
pub(crate) struct Budget {
    fields: usize,
    maximum: usize,
}

impl Budget {
    pub(crate) const fn new(limits: Limits) -> Self {
        Self {
            fields: 0,
            maximum: limits.max_fields,
        }
    }

    fn charge(&mut self) -> Result<(), Error> {
        self.fields += 1;
        if self.fields > self.maximum {
            return Err(Error::Limit {
                kind: LimitKind::Fields,
                observed: self.fields,
                maximum: self.maximum,
            });
        }
        Ok(())
    }
}

pub(crate) fn read_varint(data: &[u8], offset: &mut usize) -> Result<u64, Error> {
    let mut value = 0_u64;
    for shift in (0..64).step_by(7) {
        let Some(&byte) = data.get(*offset) else {
            return Err(Error::Malformed {
                reason: "truncated varint",
            });
        };
        *offset += 1;
        if shift == 63 && byte > 1 {
            return Err(Error::Malformed {
                reason: "varint overflows 64 bits",
            });
        }
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return Ok(value);
        }
    }
    Err(Error::Malformed {
        reason: "varint longer than ten bytes",
    })
}

pub(crate) const fn varint_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

pub(crate) fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

/// Visit every top-level field of `data` once, validating its framing.
pub(crate) fn visit_fields(
    data: &[u8],
    budget: &mut Budget,
    mut visit: impl FnMut(Field) -> Result<(), Error>,
) -> Result<(), Error> {
    let mut offset = 0;
    while offset < data.len() {
        budget.charge()?;
        let start = offset;
        let key = read_varint(data, &mut offset)?;
        let number = key >> 3;
        if number == 0 || number > MAX_FIELD_NUMBER {
            return Err(Error::Malformed {
                reason: "invalid field number",
            });
        }
        let wire_type = (key & 7) as u8;
        let key_end = offset;
        let mut varint = 0;
        let (payload_start, end) = match wire_type {
            WIRE_VARINT => {
                varint = read_varint(data, &mut offset)?;
                (key_end, offset)
            },
            WIRE_FIXED64 => (key_end, key_end.checked_add(8).unwrap_or(usize::MAX)),
            WIRE_LENGTH_DELIMITED => {
                let length = read_varint(data, &mut offset)?;
                let length = usize::try_from(length).map_err(|_| Error::Malformed {
                    reason: "length exceeds the address space",
                })?;
                (offset, offset.checked_add(length).unwrap_or(usize::MAX))
            },
            WIRE_FIXED32 => (key_end, key_end.checked_add(4).unwrap_or(usize::MAX)),
            _ => {
                return Err(Error::Malformed {
                    reason: "unsupported wire type",
                });
            },
        };
        if end > data.len() {
            return Err(Error::Malformed {
                reason: "field extends past the payload",
            });
        }
        offset = end;
        visit(Field {
            number: number as u32,
            wire_type,
            start,
            key_end,
            payload_start,
            end,
            varint,
        })?;
    }
    Ok(())
}

/// The unique occurrence of `number` in `data`, if any.
pub(crate) fn singular(
    data: &[u8],
    number: u32,
    budget: &mut Budget,
) -> Result<Option<Field>, Error> {
    let mut found = None;
    visit_fields(data, budget, |field| {
        if field.number == number && found.replace(field).is_some() {
            return Err(Error::DuplicateField { field: number });
        }
        Ok(())
    })?;
    Ok(found)
}

pub(crate) const fn require_wire_type(field: &Field, expected: u8) -> Result<(), Error> {
    if field.wire_type != expected {
        return Err(Error::WireType {
            field: field.number,
            expected,
            found: field.wire_type,
        });
    }
    Ok(())
}

pub(crate) fn check_bytes(observed: usize, limits: Limits) -> Result<(), Error> {
    if observed > limits.max_bytes {
        return Err(Error::Limit {
            kind: LimitKind::Bytes,
            observed,
            maximum: limits.max_bytes,
        });
    }
    Ok(())
}

pub(crate) fn allocate(capacity: usize, limits: Limits) -> Result<Vec<u8>, Error> {
    check_bytes(capacity, limits)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation { bytes: capacity })?;
    Ok(output)
}

/// Locate the generated extension of an outer style payload, requiring it to
/// be unique, length-delimited and canonically framed.
fn locate_extension(source: &[u8], budget: &mut Budget) -> Result<Option<Field>, Error> {
    let Some(field) = singular(source, EXTENSION_FIELD, budget)? else {
        return Ok(None);
    };
    require_wire_type(&field, WIRE_LENGTH_DELIMITED)?;
    let length = (field.end - field.payload_start) as u64;
    let canonical_key = varint_len((u64::from(EXTENSION_FIELD) << 3) | 2);
    if field.key_end - field.start != canonical_key
        || field.payload_start - field.key_end != varint_len(length)
    {
        return Err(Error::Malformed {
            reason: "the generated extension is not canonically framed",
        });
    }
    Ok(Some(field))
}

/// The generated properties of one style payload, borrowed from the source.
#[derive(Debug, Clone, Copy)]
pub struct Properties<'source> {
    archive: Archive,
    extension: Option<&'source [u8]>,
    limits: Limits,
}

/// Validate one style payload and borrow its generated extension.
///
/// The outer payload and the extension are scanned once each; every field's
/// framing is checked and charged to `limits`.
///
/// # Errors
///
/// [`Error::Malformed`], [`Error::DuplicateField`] or [`Error::WireType`] for
/// an invalid payload or extension; [`Error::Limit`] when `limits` is
/// exceeded.
pub fn decode(source: &[u8], archive: Archive, limits: Limits) -> Result<Properties<'_>, Error> {
    check_bytes(source.len(), limits)?;
    let mut budget = Budget::new(limits);
    let extension = locate_extension(source, &mut budget)?;
    let extension = extension.map(|field| field.payload(source));
    if let Some(extension) = extension {
        visit_fields(extension, &mut budget, |_| Ok(()))?;
    }
    Ok(Properties {
        archive,
        extension,
        limits,
    })
}

impl<'source> Properties<'source> {
    /// The archive these properties were decoded as.
    #[must_use]
    pub const fn archive(&self) -> Archive {
        self.archive
    }

    /// Whether the payload carries a generated extension at all.
    #[must_use]
    pub const fn has_extension(&self) -> bool {
        self.extension.is_some()
    }

    /// The raw generated extension payload.
    #[must_use]
    pub const fn extension(&self) -> Option<&'source [u8]> {
        self.extension
    }

    fn scalar(&self, number: u32, wire_type: u8) -> Result<Option<Field>, Error> {
        let Some(extension) = self.extension else {
            return Ok(None);
        };
        let mut budget = Budget::new(self.limits);
        let Some(field) = singular(extension, number, &mut budget)? else {
            return Ok(None);
        };
        require_wire_type(&field, wire_type)?;
        Ok(Some(field))
    }

    /// Whether `number` is present in the extension.
    ///
    /// # Errors
    ///
    /// [`Error::DuplicateField`] when it occurs more than once.
    pub fn contains(&self, number: u32) -> Result<bool, Error> {
        let Some(extension) = self.extension else {
            return Ok(false);
        };
        let mut budget = Budget::new(self.limits);
        Ok(singular(extension, number, &mut budget)?.is_some())
    }

    /// A `float` property.
    ///
    /// # Errors
    ///
    /// [`Error::DuplicateField`] or [`Error::WireType`].
    pub fn float(&self, number: u32) -> Result<Option<f32>, Error> {
        let Some(extension) = self.extension else {
            return Ok(None);
        };
        Ok(self
            .scalar(number, WIRE_FIXED32)?
            .map(|field| read_f32(field.payload(extension))))
    }

    /// A `bool` property (any non-zero varint is `true`, as protobuf reads it).
    ///
    /// # Errors
    ///
    /// [`Error::DuplicateField`] or [`Error::WireType`].
    pub fn bool(&self, number: u32) -> Result<Option<bool>, Error> {
        Ok(self
            .scalar(number, WIRE_VARINT)?
            .map(|field| field.varint != 0))
    }

    /// An `int32` property (the low 32 bits of the varint, as protobuf reads
    /// it).
    ///
    /// # Errors
    ///
    /// [`Error::DuplicateField`] or [`Error::WireType`].
    pub fn int32(&self, number: u32) -> Result<Option<i32>, Error> {
        Ok(self
            .scalar(number, WIRE_VARINT)?
            .map(|field| field.varint as i32))
    }

    /// The raw varint of a varint-typed property (for strict boolean or
    /// enumeration reads).
    ///
    /// # Errors
    ///
    /// [`Error::DuplicateField`] or [`Error::WireType`].
    pub fn varint(&self, number: u32) -> Result<Option<u64>, Error> {
        Ok(self.scalar(number, WIRE_VARINT)?.map(|field| field.varint))
    }

    /// The raw payload of a bytes- or string-typed property.
    ///
    /// # Errors
    ///
    /// [`Error::DuplicateField`] or [`Error::WireType`].
    pub fn bytes(&self, number: u32) -> Result<Option<&'source [u8]>, Error> {
        self.message(number)
    }

    /// The raw payload of a message-typed property (a stroke, fill or
    /// shadow), for the `shape_style_codec::native` decoders.
    ///
    /// # Errors
    ///
    /// [`Error::DuplicateField`] or [`Error::WireType`].
    pub fn message(&self, number: u32) -> Result<Option<&'source [u8]>, Error> {
        let Some(extension) = self.extension else {
            return Ok(None);
        };
        Ok(self
            .scalar(number, WIRE_LENGTH_DELIMITED)?
            .map(|field| field.payload(extension)))
    }

    /// A `TSCH.Chart3DVectorArchive` property.
    ///
    /// # Errors
    ///
    /// [`Error::DuplicateField`] or [`Error::WireType`] for the property or a
    /// component; [`Error::Malformed`] for an invalid vector payload.
    pub fn vector(&self, number: u32) -> Result<Option<Vector>, Error> {
        let Some(extension) = self.extension else {
            return Ok(None);
        };
        let Some(field) = self.scalar(number, WIRE_LENGTH_DELIMITED)? else {
            return Ok(None);
        };
        let mut budget = Budget::new(self.limits);
        decode_vector(field.payload(extension), &mut budget).map(Some)
    }
}

pub(crate) fn read_f32(payload: &[u8]) -> f32 {
    let mut bytes = [0_u8; 4];
    bytes.copy_from_slice(payload);
    f32::from_le_bytes(bytes)
}

fn decode_vector(payload: &[u8], budget: &mut Budget) -> Result<Vector, Error> {
    let mut vector = Vector::default();
    let mut seen = [false; 4];
    visit_fields(payload, budget, |component| {
        let slot = match component.number {
            field::vector::X => 0,
            field::vector::Y => 1,
            field::vector::Z => 2,
            field::vector::W => 3,
            _ => return Ok(()),
        };
        require_wire_type(&component, WIRE_FIXED32)?;
        if std::mem::replace(&mut seen[slot], true) {
            return Err(Error::DuplicateField {
                field: component.number,
            });
        }
        let value = read_f32(component.payload(payload));
        match slot {
            0 => vector.x = value,
            1 => vector.y = value,
            2 => vector.z = value,
            _ => vector.w = value,
        }
        Ok(())
    })?;
    Ok(vector)
}

/// The `TSS.StyleArchive.parent` of a style payload (field 1 of every TSCH
/// style archive; 0 reads as absent).
///
/// # Errors
///
/// Framing, duplicate-field, wire-type and limit errors.
pub fn parent(source: &[u8], limits: Limits) -> Result<Option<u64>, Error> {
    check_bytes(source.len(), limits)?;
    let mut budget = Budget::new(limits);
    let Some(style) = singular(source, 1, &mut budget)? else {
        return Ok(None);
    };
    require_wire_type(&style, WIRE_LENGTH_DELIMITED)?;
    let style = style.payload(source);
    let Some(reference) = singular(style, 3, &mut budget)? else {
        return Ok(None);
    };
    require_wire_type(&reference, WIRE_LENGTH_DELIMITED)?;
    let reference = reference.payload(style);
    let Some(identifier) = singular(reference, 1, &mut budget)? else {
        return Ok(None);
    };
    require_wire_type(&identifier, WIRE_VARINT)?;
    Ok((identifier.varint != 0).then_some(identifier.varint))
}

/// The candidate payload of a [`rewrite`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Rewrite<'source> {
    bytes: Cow<'source, [u8]>,
}

impl<'source> Rewrite<'source> {
    pub(crate) const fn from_cow(bytes: Cow<'source, [u8]>) -> Self {
        Self { bytes }
    }

    /// Whether the candidate differs from the source.
    #[must_use]
    pub fn changed(&self) -> bool {
        matches!(self.bytes, Cow::Owned(_))
    }

    /// The candidate payload.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// The candidate payload, owned.
    #[must_use]
    pub fn into_bytes(self) -> Cow<'source, [u8]> {
        self.bytes
    }
}

fn validate_write(write: Write<'_>) -> Result<(), Error> {
    let number = write.field();
    if number == 0 || number == EXTENSION_FIELD || u64::from(number) > MAX_FIELD_NUMBER {
        return Err(Error::InvalidField { field: number });
    }
    if let Write::VectorDepth { denominator, .. } = write
        && denominator != field::vector::X
        && denominator != field::vector::Y
    {
        return Err(Error::InvalidField { field: denominator });
    }
    Ok(())
}

/// Apply a batch of property writes to one style payload.
///
/// Writes apply in order. When the source has no generated extension and the
/// batch writes at least one field, the extension is created with those
/// fields in field-number order (what a protobuf encoder writes) and appended
/// to the payload; a batch that writes nothing leaves the payload unchanged.
/// An extension the batch leaves empty is kept or removed as `empty` says.
///
/// # Errors
///
/// Every [`decode`] error; [`Error::InvalidField`] for a write that names an
/// unwritable field; [`Error::InvalidScale`] for a depth over an invalid
/// stored vector; [`Error::Limit`]/[`Error::Allocation`] for the candidate.
pub fn rewrite<'source>(
    source: &'source [u8],
    archive: Archive,
    writes: &[Write<'_>],
    empty: EmptyExtension,
    limits: Limits,
) -> Result<Rewrite<'source>, Error> {
    for &write in writes {
        validate_write(write)?;
    }
    let properties = decode(source, archive, limits)?;
    let mut budget = Budget::new(limits);
    let Some(extension_field) = locate_extension(source, &mut budget)? else {
        let created = create_extension(writes, limits)?;
        let Some(created) = created else {
            return Ok(Rewrite {
                bytes: Cow::Borrowed(source),
            });
        };
        let key = (u64::from(EXTENSION_FIELD) << 3) | 2;
        let capacity =
            source.len() + varint_len(key) + varint_len(created.len() as u64) + created.len();
        let mut output = allocate(capacity, limits)?;
        output.extend_from_slice(source);
        push_varint(&mut output, key);
        push_varint(&mut output, created.len() as u64);
        output.extend_from_slice(&created);
        return Ok(Rewrite {
            bytes: Cow::Owned(output),
        });
    };
    let original = properties.extension.unwrap_or_default();
    let mut extension = Cow::Borrowed(original);
    for &write in writes {
        if let Some(next) = apply_write(&extension, write, limits)? {
            extension = Cow::Owned(next);
        }
    }
    if *extension == *original {
        return Ok(Rewrite {
            bytes: Cow::Borrowed(source),
        });
    }
    let output = if extension.is_empty() && empty == EmptyExtension::Remove {
        remove_field(source, &extension_field, limits)?
    } else {
        replace_payload(source, &extension_field, &extension, limits)?
    };
    Ok(Rewrite {
        bytes: Cow::Owned(output),
    })
}

/// A write resolved against the presence of its field.
#[derive(Debug, Clone)]
enum Resolved {
    /// Nothing to do.
    Keep,
    /// Remove the field.
    Clear,
    /// Store `value` (the encoded value bytes for the field's wire type).
    Set(Vec<u8>),
    /// Replace components of the present vector.
    Components([Option<f32>; 4]),
}

fn encode_vector(vector: Vector) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(20);
    for (number, component) in [
        (field::vector::X, vector.x),
        (field::vector::Y, vector.y),
        (field::vector::Z, vector.z),
        (field::vector::W, vector.w),
    ] {
        push_varint(
            &mut bytes,
            (u64::from(number) << 3) | u64::from(WIRE_FIXED32),
        );
        bytes.extend_from_slice(&component.to_le_bytes());
    }
    bytes
}

fn int32_bytes(value: i32) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(10);
    push_varint(&mut bytes, i64::from(value) as u64);
    bytes
}

/// Resolve a write; `present` is the stored vector payload (for vector
/// writes) or an empty slice for any other present field.
fn resolve(
    write: Write<'_>,
    present: Option<&[u8]>,
    budget: &mut Budget,
    limits: Limits,
) -> Result<Resolved, Error> {
    let set_or_clear = |value: Option<Vec<u8>>| value.map_or(Resolved::Clear, Resolved::Set);
    Ok(match write {
        Write::Float { value, .. } => set_or_clear(value.map(|v| v.to_le_bytes().to_vec())),
        Write::Bool { value, .. } => set_or_clear(value.map(|v| vec![u8::from(v)])),
        Write::Int32 { value, .. } => set_or_clear(value.map(int32_bytes)),
        Write::Vector { value, .. } => set_or_clear(value.map(encode_vector)),
        Write::Message { value, .. } => set_or_clear(value.map(<[u8]>::to_vec)),
        Write::MergeMessage { value, schema, .. } => Resolved::Set(match present {
            Some(existing) => merge::merge(existing, value, schema, limits)?,
            None => {
                merge::validate(value, schema, limits)?;
                value.to_vec()
            },
        }),
        Write::FloatUnlessAbsentDefault { value, default, .. } => {
            if present.is_none() && value == default {
                Resolved::Keep
            } else {
                Resolved::Set(value.to_le_bytes().to_vec())
            }
        },
        Write::BoolUnlessAbsentDefault { value, default, .. } => {
            if present.is_none() && value == default {
                Resolved::Keep
            } else {
                Resolved::Set(vec![u8::from(value)])
            }
        },
        Write::VectorComponents {
            components, absent, ..
        } => match present {
            Some(_) => Resolved::Components(components),
            None => absent.map_or(Resolved::Keep, |vector| {
                Resolved::Set(encode_vector(vector))
            }),
        },
        Write::VectorDepth {
            field: number,
            ratio,
            denominator,
            absent,
        } => match present {
            Some(stored) => {
                let stored = decode_vector(stored, budget)?;
                let denominator = if denominator == field::vector::Y {
                    stored.y
                } else {
                    stored.x
                };
                if !denominator.is_finite() || denominator <= 0.0 {
                    return Err(Error::InvalidScale { field: number });
                }
                Resolved::Components([None, None, Some(ratio * denominator), None])
            },
            None => absent.map_or(Resolved::Keep, |vector| {
                Resolved::Set(encode_vector(vector))
            }),
        },
    })
}

fn create_extension(writes: &[Write<'_>], limits: Limits) -> Result<Option<Vec<u8>>, Error> {
    // Sequential application over an empty extension: later writes of the
    // same field see the earlier ones.
    let mut fields: Vec<(u32, u8, Vec<u8>)> = Vec::new();
    let mut budget = Budget::new(limits);
    for &write in writes {
        let number = write.field();
        let index = fields.iter().position(|(kept, ..)| *kept == number);
        let present = index.map(|index| fields[index].2.as_slice());
        match resolve(write, present, &mut budget, limits)? {
            Resolved::Keep => {},
            Resolved::Clear => {
                if let Some(index) = index {
                    fields.remove(index);
                }
            },
            Resolved::Set(value) => match index {
                Some(index) => fields[index].2 = value,
                None => fields.push((number, write.wire_type(), value)),
            },
            Resolved::Components(components) => {
                let Some(index) = index else {
                    continue;
                };
                let mut vector = fields[index].2.clone();
                for (component, value) in (field::vector::X..).zip(components) {
                    if let Some(value) = value
                        && let Some(next) = apply_component(&vector, component, value, limits)?
                    {
                        vector = next;
                    }
                }
                fields[index].2 = vector;
            },
        }
    }
    if fields.is_empty() {
        return Ok(None);
    }
    fields.sort_by_key(|(number, ..)| *number);
    let mut output = Vec::new();
    for (number, wire_type, value) in fields {
        push_field(&mut output, number, wire_type, &value);
        check_bytes(output.len(), limits)?;
    }
    Ok(Some(output))
}

fn push_field(output: &mut Vec<u8>, number: u32, wire_type: u8, value: &[u8]) {
    push_varint(output, (u64::from(number) << 3) | u64::from(wire_type));
    if wire_type == WIRE_LENGTH_DELIMITED {
        push_varint(output, value.len() as u64);
    }
    output.extend_from_slice(value);
}

/// Append one whole field for a value-carrying plain write (tests build
/// payloads with it).
#[cfg(test)]
fn encode_field(output: &mut Vec<u8>, write: Write<'_>) {
    let mut budget = Budget::new(Limits::DEFAULT);
    if let Ok(Resolved::Set(value)) = resolve(write, None, &mut budget, Limits::DEFAULT) {
        push_field(output, write.field(), write.wire_type(), &value);
    }
}

/// Apply one write to an extension payload; `None` when the bytes are
/// unchanged.
fn apply_write(
    extension: &[u8],
    write: Write<'_>,
    limits: Limits,
) -> Result<Option<Vec<u8>>, Error> {
    let mut budget = Budget::new(limits);
    let existing = singular(extension, write.field(), &mut budget)?;
    if let Some(field) = &existing {
        require_wire_type(field, write.wire_type())?;
    }
    let present = existing.map(|field| field.payload(extension));
    let resolved = resolve(write, present, &mut budget, limits)?;
    match (resolved, existing) {
        (Resolved::Keep, _) | (Resolved::Clear | Resolved::Components(_), None) => Ok(None),
        (Resolved::Clear, Some(field)) => remove_field(extension, &field, limits).map(Some),
        (Resolved::Set(value), None) => {
            let mut output = allocate(extension.len() + value.len() + 15, limits)?;
            output.extend_from_slice(extension);
            push_field(&mut output, write.field(), write.wire_type(), &value);
            check_bytes(output.len(), limits)?;
            Ok(Some(output))
        },
        (Resolved::Set(value), Some(field)) => {
            if write.wire_type() == WIRE_LENGTH_DELIMITED {
                if field.payload(extension) == value.as_slice() {
                    return Ok(None);
                }
                replace_payload(extension, &field, &value, limits).map(Some)
            } else {
                replace_scalar(extension, &field, &value, limits)
            }
        },
        (Resolved::Components(components), Some(field)) => {
            let mut vector = Cow::Borrowed(field.payload(extension));
            for (component, value) in (field::vector::X..).zip(components) {
                if let Some(value) = value
                    && let Some(next) = apply_component(&vector, component, value, limits)?
                {
                    vector = Cow::Owned(next);
                }
            }
            if *vector == *field.payload(extension) {
                return Ok(None);
            }
            replace_payload(extension, &field, &vector, limits).map(Some)
        },
    }
}

/// Set one fixed32 component of a vector payload.
fn apply_component(
    vector: &[u8],
    component: u32,
    value: f32,
    limits: Limits,
) -> Result<Option<Vec<u8>>, Error> {
    let mut budget = Budget::new(limits);
    let bytes = value.to_le_bytes();
    match singular(vector, component, &mut budget)? {
        Some(field) => {
            require_wire_type(&field, WIRE_FIXED32)?;
            replace_scalar(vector, &field, &bytes, limits)
        },
        None => {
            let mut output = allocate(vector.len() + 5, limits)?;
            output.extend_from_slice(vector);
            push_field(&mut output, component, WIRE_FIXED32, &bytes);
            Ok(Some(output))
        },
    }
}

pub(crate) fn replace_scalar(
    data: &[u8],
    field: &Field,
    value: &[u8],
    limits: Limits,
) -> Result<Option<Vec<u8>>, Error> {
    if &data[field.key_end..field.end] == value {
        return Ok(None);
    }
    let capacity = data.len() - (field.end - field.key_end) + value.len();
    let mut output = allocate(capacity, limits)?;
    output.extend_from_slice(&data[..field.key_end]);
    output.extend_from_slice(value);
    output.extend_from_slice(&data[field.end..]);
    Ok(Some(output))
}

/// Replace a length-delimited field's payload, keeping its key bytes and,
/// when the length is unchanged, its length prefix.
pub(crate) fn replace_payload(
    data: &[u8],
    field: &Field,
    payload: &[u8],
    limits: Limits,
) -> Result<Vec<u8>, Error> {
    let old_length = field.end - field.payload_start;
    let keep_prefix = old_length == payload.len();
    let prefix_length = if keep_prefix {
        field.payload_start - field.key_end
    } else {
        varint_len(payload.len() as u64)
    };
    let capacity = data.len() - (field.end - field.key_end) + prefix_length + payload.len();
    let mut output = allocate(capacity, limits)?;
    output.extend_from_slice(&data[..field.key_end]);
    if keep_prefix {
        output.extend_from_slice(&data[field.key_end..field.payload_start]);
    } else {
        push_varint(&mut output, payload.len() as u64);
    }
    output.extend_from_slice(payload);
    output.extend_from_slice(&data[field.end..]);
    Ok(output)
}

pub(crate) fn remove_field(data: &[u8], field: &Field, limits: Limits) -> Result<Vec<u8>, Error> {
    let mut output = allocate(data.len() - (field.end - field.start), limits)?;
    output.extend_from_slice(&data[..field.start]);
    output.extend_from_slice(&data[field.end..]);
    Ok(output)
}

/// The native write rules of the inspector properties the shared chart
/// engine edits: one call maps a property's stored value to the archive it
/// lives in, the batch of [`Write`]s that stores it and the empty-extension
/// policy (the migration host's per-property rules, record 0375). Format
/// families convert their `litchi_iwa_common` value to the stored natives
/// (`degrees()`, `native_fraction()`, `native_ratio()`, …) and pass the
/// result to [`rewrite`] unchanged.
pub mod property {
    use super::{Archive, EmptyExtension, Vector, Write, field};

    const MAX_WRITES: usize = 3;
    const NO_WRITE: Write<'static> = Write::Float {
        field: field::chart_non_style::PIE_START_ANGLE,
        value: None,
    };
    /// Stored donut inner radius when the property is absent (75%).
    pub const DEFAULT_DONUT_INNER_RADIUS: f32 = 75.0 / 100.0;
    /// Stored 3D series-gap ratio when the property is absent (100%).
    pub const DEFAULT_SERIES_GAP_3D: f32 = 1.0;
    /// Rotation `(x, y)` degrees when no rotation vector is stored.
    pub const DEFAULT_ROTATION_3D: (f32, f32) = (18.5, -18.125);
    /// Stored per-wedge label distance (percent) when no override exists.
    pub const DEFAULT_PIE_LABEL_DISTANCE: f32 = 67.0;

    /// The archive, writes and empty-extension policy of one property value.
    #[derive(Debug, Clone, Copy, PartialEq)]
    pub struct PropertyWrite<'a> {
        archive: Archive,
        writes: [Write<'a>; MAX_WRITES],
        len: usize,
        empty: EmptyExtension,
    }

    impl<'a> PropertyWrite<'a> {
        const fn one(archive: Archive, write: Write<'a>, empty: EmptyExtension) -> Self {
            Self {
                archive,
                writes: [write, NO_WRITE, NO_WRITE],
                len: 1,
                empty,
            }
        }

        const fn two(
            archive: Archive,
            first: Write<'a>,
            second: Write<'a>,
            empty: EmptyExtension,
        ) -> Self {
            Self {
                archive,
                writes: [first, second, NO_WRITE],
                len: 2,
                empty,
            }
        }

        const fn three(archive: Archive, writes: [Write<'a>; 3], empty: EmptyExtension) -> Self {
            Self {
                archive,
                writes,
                len: 3,
                empty,
            }
        }

        /// The style archive that stores the property.
        #[must_use]
        pub const fn archive(&self) -> Archive {
            self.archive
        }

        /// The writes, in order.
        #[must_use]
        pub fn writes(&self) -> &[Write<'a>] {
            &self.writes[..self.len]
        }

        /// The policy for an extension the writes leave empty.
        #[must_use]
        pub const fn empty(&self) -> EmptyExtension {
            self.empty
        }
    }

    /// The per-kind scale vector that stores a 3D chart's depth.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    #[non_exhaustive]
    pub enum Scale3d {
        /// `tschchartinfoarea3dscale`.
        Area,
        /// `tschchartinfobar3dscale` (horizontal: the ratio is relative to `y`).
        Bar,
        /// `tschchartinfocolumn3dscale`.
        Column,
        /// `tschchartinfoline3dscale`.
        Line,
        /// `tschchartinfopie3dscale`.
        Pie,
    }

    impl Scale3d {
        /// The extension field of the scale vector.
        #[must_use]
        pub const fn field(self) -> u32 {
            match self {
                Self::Area => field::chart_non_style::AREA_SCALE_3D,
                Self::Bar => field::chart_non_style::BAR_SCALE_3D,
                Self::Column => field::chart_non_style::COLUMN_SCALE_3D,
                Self::Line => field::chart_non_style::LINE_SCALE_3D,
                Self::Pie => field::chart_non_style::PIE_SCALE_3D,
            }
        }

        const fn denominator(self) -> u32 {
            match self {
                Self::Bar => field::vector::Y,
                _ => field::vector::X,
            }
        }
    }

    const fn chart_float(field: u32, value: f32, default: f32) -> PropertyWrite<'static> {
        PropertyWrite::one(
            Archive::ChartNonStyle,
            Write::Float {
                field,
                value: if value == default { None } else { Some(value) },
            },
            EmptyExtension::Keep,
        )
    }

    /// Pie/donut start angle (stored degrees; 0 clears).
    #[must_use]
    pub const fn pie_start_angle(degrees: f32) -> PropertyWrite<'static> {
        chart_float(field::chart_non_style::PIE_START_ANGLE, degrees, 0.0)
    }

    /// Donut inner radius (stored fraction; the default clears).
    #[must_use]
    pub const fn donut_inner_radius(fraction: f32) -> PropertyWrite<'static> {
        chart_float(
            field::chart_non_style::DONUT_INNER_RADIUS,
            fraction,
            DEFAULT_DONUT_INNER_RADIUS,
        )
    }

    /// Radar start angle (stored degrees; 0 clears).
    #[must_use]
    pub const fn radar_start_angle(degrees: f32) -> PropertyWrite<'static> {
        chart_float(field::chart_non_style::RADAR_START_ANGLE, degrees, 0.0)
    }

    /// 3D series gap (stored ratio; the default clears).
    #[must_use]
    pub const fn series_gap_3d(ratio: f32) -> PropertyWrite<'static> {
        chart_float(
            field::chart_non_style::SERIES_GAP_3D,
            ratio,
            DEFAULT_SERIES_GAP_3D,
        )
    }

    /// 3D scene rotation in degrees: the default clears the vector; otherwise
    /// a stored vector keeps `z`/`w` and an absent one is written as
    /// `(x, y, 0, 0)`.
    #[must_use]
    pub const fn rotation_3d(x_degrees: f32, y_degrees: f32) -> PropertyWrite<'static> {
        let field = field::chart_non_style::ROTATION_3D;
        let write = if x_degrees == DEFAULT_ROTATION_3D.0 && y_degrees == DEFAULT_ROTATION_3D.1 {
            Write::Vector { field, value: None }
        } else {
            Write::VectorComponents {
                field,
                components: [Some(x_degrees), Some(y_degrees), None, None],
                absent: Some(Vector {
                    x: x_degrees,
                    y: y_degrees,
                    z: 0.0,
                    w: 0.0,
                }),
            }
        };
        PropertyWrite::one(Archive::ChartNonStyle, write, EmptyExtension::Keep)
    }

    /// 3D depth: a stored scale vector gets `z = ratio × denominator`; an
    /// absent one is written as `absent`, or left absent when `absent` is
    /// `None` (the caller passes `None` for the default depth).
    #[must_use]
    pub const fn depth_3d(
        scale: Scale3d,
        ratio: f32,
        absent: Option<[f32; 4]>,
    ) -> PropertyWrite<'static> {
        let absent = match absent {
            Some([x, y, z, w]) => Some(Vector { x, y, z, w }),
            None => None,
        };
        PropertyWrite::one(
            Archive::ChartNonStyle,
            Write::VectorDepth {
                field: scale.field(),
                ratio,
                denominator: scale.denominator(),
                absent,
            },
            EmptyExtension::Keep,
        )
    }

    /// Rounded corners (stored radius fraction, outside-corners switch): the
    /// default (0, false) clears both; otherwise a default-valued field is
    /// written only when it is already present.
    #[must_use]
    pub const fn rounded_corners(
        radius: f32,
        outside_corners_only: bool,
    ) -> PropertyWrite<'static> {
        let radius_field = field::chart_style::ROUNDED_CORNER_RADIUS;
        let flag_field = field::chart_style::ROUNDED_CORNER_OUTER_END_ONLY;
        if radius == 0.0 && !outside_corners_only {
            return PropertyWrite::two(
                Archive::ChartStyle,
                Write::Float {
                    field: radius_field,
                    value: None,
                },
                Write::Bool {
                    field: flag_field,
                    value: None,
                },
                EmptyExtension::Keep,
            );
        }
        PropertyWrite::two(
            Archive::ChartStyle,
            Write::FloatUnlessAbsentDefault {
                field: radius_field,
                value: radius,
                default: 0.0,
            },
            Write::BoolUnlessAbsentDefault {
                field: flag_field,
                value: outside_corners_only,
                default: false,
            },
            EmptyExtension::Keep,
        )
    }

    /// A series stroke (series style): `Some(encoded)` merges the encoded
    /// `TSD.StrokeArchive` into a stored stroke (or sets it), `None` removes
    /// the local override; an emptied extension is removed. `field` is the
    /// family's field (`litchi_iwa_common::chart::series::StrokeFamily`).
    #[must_use]
    pub const fn series_stroke(field: u32, encoded: Option<&[u8]>) -> PropertyWrite<'_> {
        let write = match encoded {
            Some(value) => Write::MergeMessage {
                field,
                value,
                schema: super::merge::schema::STROKE,
            },
            None => Write::Message { field, value: None },
        };
        PropertyWrite::one(Archive::SeriesStyle, write, EmptyExtension::Remove)
    }

    /// A series fill or a data-symbol outline (series style): `Some` replaces
    /// the stored message with the encoded `TSD.FillArchive` or
    /// `TSD.StrokeArchive`, `None` removes it; an emptied extension is
    /// removed.
    #[must_use]
    pub const fn series_message(field: u32, encoded: Option<&[u8]>) -> PropertyWrite<'_> {
        PropertyWrite::one(
            Archive::SeriesStyle,
            Write::Message {
                field,
                value: encoded,
            },
            EmptyExtension::Remove,
        )
    }

    /// A data-symbol size in points (series style; `None` is automatic).
    #[must_use]
    pub const fn symbol_size(field: u32, points: Option<f32>) -> PropertyWrite<'static> {
        PropertyWrite::one(
            Archive::SeriesStyle,
            Write::Float {
                field,
                value: points,
            },
            EmptyExtension::Remove,
        )
    }

    /// What a data-symbol fill stores.
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub enum SymbolFill<'a> {
        /// An encoded `TSD.FillArchive`.
        Custom(&'a [u8]),
        /// Follow the series fill.
        SeriesFill,
        /// Follow the series stroke.
        SeriesStroke,
        /// Remove the local override.
        Reset,
    }

    /// A data-symbol fill (series style): the fill message and the two
    /// switches, in that order. `None` when `value` follows the series fill
    /// but the family has no such switch.
    #[must_use]
    pub const fn symbol_fill(
        fill_field: u32,
        use_series_fill_field: Option<u32>,
        use_series_stroke_field: u32,
        value: SymbolFill<'_>,
    ) -> Option<PropertyWrite<'_>> {
        let fill = Write::Message {
            field: fill_field,
            value: match value {
                SymbolFill::Custom(bytes) => Some(bytes),
                SymbolFill::SeriesFill | SymbolFill::SeriesStroke | SymbolFill::Reset => None,
            },
        };
        let use_stroke = Write::Bool {
            field: use_series_stroke_field,
            value: if matches!(value, SymbolFill::SeriesStroke) {
                Some(true)
            } else {
                None
            },
        };
        match use_series_fill_field {
            Some(field) => Some(PropertyWrite::three(
                Archive::SeriesStyle,
                [
                    fill,
                    Write::Bool {
                        field,
                        value: if matches!(value, SymbolFill::SeriesFill) {
                            Some(true)
                        } else {
                            None
                        },
                    },
                    use_stroke,
                ],
                EmptyExtension::Remove,
            )),
            None if matches!(value, SymbolFill::SeriesFill) => None,
            None => Some(PropertyWrite::two(
                Archive::SeriesStyle,
                fill,
                use_stroke,
                EmptyExtension::Remove,
            )),
        }
    }

    /// One wedge's explosion (stored fraction; 0 clears and an emptied
    /// extension is removed).
    #[must_use]
    pub const fn pie_wedge_explosion(fraction: f32) -> PropertyWrite<'static> {
        PropertyWrite::one(
            Archive::SeriesNonStyle,
            Write::Float {
                field: field::series_non_style::PIE_WEDGE_EXPLOSION,
                value: if fraction == 0.0 {
                    None
                } else {
                    Some(fraction)
                },
            },
            EmptyExtension::Remove,
        )
    }

    /// One wedge's label distance (stored percent, in the legacy and the
    /// current field; the default clears both and an emptied extension is
    /// removed).
    #[must_use]
    pub const fn pie_label_distance(percent: f32) -> PropertyWrite<'static> {
        let value = if percent == DEFAULT_PIE_LABEL_DISTANCE {
            None
        } else {
            Some(percent)
        };
        PropertyWrite::two(
            Archive::SeriesNonStyle,
            Write::Float {
                field: field::series_non_style::PIE_LEGACY_LABEL_EXPLOSION,
                value,
            },
            Write::Float {
                field: field::series_non_style::PIE_LABEL_EXPLOSION,
                value,
            },
            EmptyExtension::Remove,
        )
    }
}

/// A `TSP.SparseReferenceArray`: a series count and the present entries.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct SparseReferences {
    count: u32,
    entries: Vec<(u32, u64)>,
}

impl SparseReferences {
    /// The array length (`count`).
    #[must_use]
    pub const fn count(&self) -> u32 {
        self.count
    }

    /// `(index, object identifier)` of every present entry, in source order.
    #[must_use]
    pub fn entries(&self) -> &[(u32, u64)] {
        &self.entries
    }

    /// The object at `index`, if present.
    #[must_use]
    pub fn get(&self, index: u32) -> Option<u64> {
        self.entries
            .iter()
            .find(|(entry, _)| *entry == index)
            .map(|(_, identifier)| *identifier)
    }
}

/// The style and paragraph-style references of one chart
/// (`TSCH.ChartArchive`, extension 10000 of `TSCH.ChartDrawableArchive`).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct StyleReferences {
    chart_style: Option<u64>,
    chart_non_style: Option<u64>,
    legend_style: Option<u64>,
    legend_non_style: Option<u64>,
    value_axis_styles: Vec<u64>,
    value_axis_non_styles: Vec<u64>,
    category_axis_styles: Vec<u64>,
    category_axis_non_styles: Vec<u64>,
    series_theme_styles: Vec<u64>,
    series_private_styles: Option<SparseReferences>,
    series_non_styles: Option<SparseReferences>,
    paragraph_styles: Vec<u64>,
}

impl StyleReferences {
    /// `chart_style` (field 9).
    #[must_use]
    pub const fn chart_style(&self) -> Option<u64> {
        self.chart_style
    }

    /// `chart_non_style` (field 10).
    #[must_use]
    pub const fn chart_non_style(&self) -> Option<u64> {
        self.chart_non_style
    }

    /// `legend_style` (field 11).
    #[must_use]
    pub const fn legend_style(&self) -> Option<u64> {
        self.legend_style
    }

    /// `legend_non_style` (field 12).
    #[must_use]
    pub const fn legend_non_style(&self) -> Option<u64> {
        self.legend_non_style
    }

    /// `value_axis_styles` (field 13).
    #[must_use]
    pub fn value_axis_styles(&self) -> &[u64] {
        &self.value_axis_styles
    }

    /// `value_axis_nonstyles` (field 14).
    #[must_use]
    pub fn value_axis_non_styles(&self) -> &[u64] {
        &self.value_axis_non_styles
    }

    /// `category_axis_styles` (field 15).
    #[must_use]
    pub fn category_axis_styles(&self) -> &[u64] {
        &self.category_axis_styles
    }

    /// `category_axis_nonstyles` (field 16).
    #[must_use]
    pub fn category_axis_non_styles(&self) -> &[u64] {
        &self.category_axis_non_styles
    }

    /// `series_theme_styles` (field 17).
    #[must_use]
    pub fn series_theme_styles(&self) -> &[u64] {
        &self.series_theme_styles
    }

    /// `series_private_styles` (field 18).
    #[must_use]
    pub const fn series_private_styles(&self) -> Option<&SparseReferences> {
        self.series_private_styles.as_ref()
    }

    /// `series_non_styles` (field 19).
    #[must_use]
    pub const fn series_non_styles(&self) -> Option<&SparseReferences> {
        self.series_non_styles.as_ref()
    }

    /// `paragraph_styles` (field 20).
    #[must_use]
    pub fn paragraph_styles(&self) -> &[u64] {
        &self.paragraph_styles
    }

    /// The effective series style of series `index`: the private override of
    /// `series_private_styles`, else the theme style.
    #[must_use]
    pub fn series_style(&self, index: usize) -> Option<u64> {
        let private = u32::try_from(index)
            .ok()
            .and_then(|index| self.series_private_styles.as_ref()?.get(index));
        private.or_else(|| self.series_theme_styles.get(index).copied())
    }

    /// The object that stores `archive` for this chart, for the chart- and
    /// legend-level archives (`None` for the per-axis and per-series ones).
    #[must_use]
    pub const fn chart_level(&self, archive: Archive) -> Option<u64> {
        match archive {
            Archive::ChartStyle => self.chart_style,
            Archive::ChartNonStyle => self.chart_non_style,
            Archive::LegendStyle => self.legend_style,
            Archive::LegendNonStyle => self.legend_non_style,
            Archive::AxisStyle
            | Archive::AxisNonStyle
            | Archive::SeriesStyle
            | Archive::SeriesNonStyle => None,
        }
    }
}

const CHART_STYLE_FIELD: u32 = 9;
const CHART_NON_STYLE_FIELD: u32 = 10;
const LEGEND_STYLE_FIELD: u32 = 11;
const LEGEND_NON_STYLE_FIELD: u32 = 12;
const VALUE_AXIS_STYLES_FIELD: u32 = 13;
const VALUE_AXIS_NON_STYLES_FIELD: u32 = 14;
const CATEGORY_AXIS_STYLES_FIELD: u32 = 15;
const CATEGORY_AXIS_NON_STYLES_FIELD: u32 = 16;
const SERIES_THEME_STYLES_FIELD: u32 = 17;
const SERIES_PRIVATE_STYLES_FIELD: u32 = 18;
const SERIES_NON_STYLES_FIELD: u32 = 19;
const PARAGRAPH_STYLES_FIELD: u32 = 20;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const SPARSE_COUNT_FIELD: u32 = 1;
const SPARSE_ENTRIES_FIELD: u32 = 2;
const SPARSE_ENTRY_INDEX_FIELD: u32 = 1;
const SPARSE_ENTRY_REFERENCE_FIELD: u32 = 2;

/// Project the style references of one `TSCH.ChartDrawableArchive` payload.
///
/// Only the chart extension's top-level fields and the reference messages are
/// visited; the data grid and every other field are skipped by length. A
/// reference with identifier 0 reads as absent.
///
/// # Errors
///
/// [`Error::Malformed`] for invalid framing or a chart drawable without a
/// chart extension, [`Error::DuplicateField`]/[`Error::WireType`] for an
/// invalid reference, [`Error::Limit`] when `limits` is exceeded.
pub fn references(chart_drawable: &[u8], limits: Limits) -> Result<StyleReferences, Error> {
    check_bytes(chart_drawable.len(), limits)?;
    let mut budget = Budget::new(limits);
    let chart =
        singular(chart_drawable, EXTENSION_FIELD, &mut budget)?.ok_or(Error::Malformed {
            reason: "the chart drawable has no chart extension",
        })?;
    require_wire_type(&chart, WIRE_LENGTH_DELIMITED)?;
    let chart = chart.payload(chart_drawable);
    let mut references = StyleReferences::default();
    let mut singular_seen = [false; 4];
    let mut sparse_seen = [false; 2];
    let mut nested = Budget::new(limits);
    visit_fields(chart, &mut budget, |field| {
        match field.number {
            CHART_STYLE_FIELD..=LEGEND_NON_STYLE_FIELD => {
                let slot = (field.number - CHART_STYLE_FIELD) as usize;
                if std::mem::replace(&mut singular_seen[slot], true) {
                    return Err(Error::DuplicateField {
                        field: field.number,
                    });
                }
                let identifier = reference(chart, &field, &mut nested)?;
                let target = match field.number {
                    CHART_STYLE_FIELD => &mut references.chart_style,
                    CHART_NON_STYLE_FIELD => &mut references.chart_non_style,
                    LEGEND_STYLE_FIELD => &mut references.legend_style,
                    _ => &mut references.legend_non_style,
                };
                *target = identifier;
            },
            VALUE_AXIS_STYLES_FIELD..=SERIES_THEME_STYLES_FIELD | PARAGRAPH_STYLES_FIELD => {
                let identifier = reference(chart, &field, &mut nested)?.unwrap_or(0);
                let target = match field.number {
                    VALUE_AXIS_STYLES_FIELD => &mut references.value_axis_styles,
                    VALUE_AXIS_NON_STYLES_FIELD => &mut references.value_axis_non_styles,
                    CATEGORY_AXIS_STYLES_FIELD => &mut references.category_axis_styles,
                    CATEGORY_AXIS_NON_STYLES_FIELD => &mut references.category_axis_non_styles,
                    SERIES_THEME_STYLES_FIELD => &mut references.series_theme_styles,
                    _ => &mut references.paragraph_styles,
                };
                target.push(identifier);
            },
            SERIES_PRIVATE_STYLES_FIELD | SERIES_NON_STYLES_FIELD => {
                let slot = (field.number - SERIES_PRIVATE_STYLES_FIELD) as usize;
                if std::mem::replace(&mut sparse_seen[slot], true) {
                    return Err(Error::DuplicateField {
                        field: field.number,
                    });
                }
                require_wire_type(&field, WIRE_LENGTH_DELIMITED)?;
                let sparse = sparse_references(field.payload(chart), &mut nested)?;
                if field.number == SERIES_PRIVATE_STYLES_FIELD {
                    references.series_private_styles = Some(sparse);
                } else {
                    references.series_non_styles = Some(sparse);
                }
            },
            _ => {},
        }
        Ok(())
    })?;
    Ok(references)
}

const CHART_LEGEND_FRAME_FIELD: u32 = 3;

fn chart_extension(chart_drawable: &[u8], budget: &mut Budget) -> Result<Field, Error> {
    let chart = singular(chart_drawable, EXTENSION_FIELD, budget)?.ok_or(Error::Malformed {
        reason: "the chart drawable has no chart extension",
    })?;
    require_wire_type(&chart, WIRE_LENGTH_DELIMITED)?;
    Ok(chart)
}

fn exactly_one_float(data: &[u8], number: u32, budget: &mut Budget) -> Result<f32, Error> {
    let field = singular(data, number, budget)?.ok_or(Error::Malformed {
        reason: "a legend-frame coordinate is missing",
    })?;
    require_wire_type(&field, WIRE_FIXED32)?;
    Ok(read_f32(field.payload(data)))
}

fn exactly_one_message<'a>(
    data: &'a [u8],
    number: u32,
    budget: &mut Budget,
) -> Result<&'a [u8], Error> {
    let field = singular(data, number, budget)?.ok_or(Error::Malformed {
        reason: "a legend-frame origin or size is missing",
    })?;
    require_wire_type(&field, WIRE_LENGTH_DELIMITED)?;
    Ok(field.payload(data))
}

/// The explicit legend rectangle `[x, y, width, height]` of a chart
/// (`TSCH.ChartArchive.legend_frame`), or `None` for automatic placement.
/// Every component must be stored exactly once.
///
/// # Errors
///
/// [`Error::Malformed`] for a chart drawable without a chart extension or an
/// incomplete rectangle; framing, duplicate-field, wire-type and limit
/// errors.
pub fn legend_frame(chart_drawable: &[u8], limits: Limits) -> Result<Option<[f32; 4]>, Error> {
    check_bytes(chart_drawable.len(), limits)?;
    let mut budget = Budget::new(limits);
    let chart = chart_extension(chart_drawable, &mut budget)?;
    let chart = chart.payload(chart_drawable);
    let Some(frame) = singular(chart, CHART_LEGEND_FRAME_FIELD, &mut budget)? else {
        return Ok(None);
    };
    require_wire_type(&frame, WIRE_LENGTH_DELIMITED)?;
    let frame = frame.payload(chart);
    let origin = exactly_one_message(frame, 1, &mut budget)?;
    let size = exactly_one_message(frame, 2, &mut budget)?;
    Ok(Some([
        exactly_one_float(origin, 1, &mut budget)?,
        exactly_one_float(origin, 2, &mut budget)?,
        exactly_one_float(size, 1, &mut budget)?,
        exactly_one_float(size, 2, &mut budget)?,
    ]))
}

/// Set (`Some([x, y, width, height])`) or remove (`None`) a chart's legend
/// rectangle, replacing the stored field in place or appending it to the
/// chart archive and keeping every other byte.
///
/// # Errors
///
/// The [`legend_frame`] errors; [`Error::Limit`]/[`Error::Allocation`] for
/// the candidate.
pub fn rewrite_legend_frame(
    chart_drawable: &[u8],
    frame: Option<[f32; 4]>,
    limits: Limits,
) -> Result<Rewrite<'_>, Error> {
    legend_frame(chart_drawable, limits)?;
    let mut budget = Budget::new(limits);
    let chart_field = chart_extension(chart_drawable, &mut budget)?;
    let chart = chart_field.payload(chart_drawable);
    let existing = singular(chart, CHART_LEGEND_FRAME_FIELD, &mut budget)?;
    let encoded = frame.map(|[x, y, width, height]| {
        let pair = |first: f32, second: f32| {
            let mut bytes = Vec::with_capacity(10);
            bytes.push(0x0d);
            bytes.extend_from_slice(&first.to_le_bytes());
            bytes.push(0x15);
            bytes.extend_from_slice(&second.to_le_bytes());
            bytes
        };
        let mut rect = Vec::with_capacity(24);
        for part in [pair(x, y), pair(width, height)].iter().enumerate() {
            rect.push(if part.0 == 0 { 0x0a } else { 0x12 });
            push_varint(&mut rect, part.1.len() as u64);
            rect.extend_from_slice(part.1);
        }
        rect
    });
    let next_chart = match (existing, encoded) {
        (None, None) => None,
        (Some(field), None) => Some(remove_field(chart, &field, limits)?),
        (Some(field), Some(rect)) => {
            if field.payload(chart) == rect.as_slice() {
                None
            } else {
                Some(replace_payload(chart, &field, &rect, limits)?)
            }
        },
        (None, Some(rect)) => {
            let mut output = allocate(chart.len() + rect.len() + 2, limits)?;
            output.extend_from_slice(chart);
            push_varint(&mut output, (u64::from(CHART_LEGEND_FRAME_FIELD) << 3) | 2);
            push_varint(&mut output, rect.len() as u64);
            output.extend_from_slice(&rect);
            Some(output)
        },
    };
    let Some(next_chart) = next_chart else {
        return Ok(Rewrite {
            bytes: Cow::Borrowed(chart_drawable),
        });
    };
    let output = replace_payload(chart_drawable, &chart_field, &next_chart, limits)?;
    Ok(Rewrite {
        bytes: Cow::Owned(output),
    })
}

/// Replace a chart's `series_non_styles` sparse array (field 19 of the chart
/// archive) with `entries` (`(series index, object identifier)` in order),
/// encoded as a protobuf encoder writes it: the `count` (the number of
/// entries) and each entry's index and reference identifier. A stored array
/// is replaced in place; an absent one is inserted before the first chart
/// field numbered above 19. Every other byte is kept.
///
/// # Errors
///
/// The [`references`] errors; [`Error::InvalidField`] for a zero
/// identifier; [`Error::Limit`]/[`Error::Allocation`] for the candidate.
pub fn rewrite_series_non_styles<'source>(
    chart_drawable: &'source [u8],
    entries: &[(u32, u64)],
    limits: Limits,
) -> Result<Rewrite<'source>, Error> {
    references(chart_drawable, limits)?;
    if entries.iter().any(|(_, identifier)| *identifier == 0) {
        return Err(Error::InvalidField {
            field: SERIES_NON_STYLES_FIELD,
        });
    }
    let mut sparse = allocate(entries.len() * 16 + 6, limits)?;
    sparse.push(0x08);
    push_varint(&mut sparse, entries.len() as u64);
    for &(index, identifier) in entries {
        let mut reference = Vec::with_capacity(11);
        reference.push(0x08);
        push_varint(&mut reference, identifier);
        let mut entry = Vec::with_capacity(reference.len() + 8);
        entry.push(0x08);
        push_varint(&mut entry, u64::from(index));
        entry.push(0x12);
        push_varint(&mut entry, reference.len() as u64);
        entry.extend_from_slice(&reference);
        sparse.push(0x12);
        push_varint(&mut sparse, entry.len() as u64);
        sparse.extend_from_slice(&entry);
    }
    let mut budget = Budget::new(limits);
    let chart_field = chart_extension(chart_drawable, &mut budget)?;
    let chart = chart_field.payload(chart_drawable);
    let mut existing = None;
    let mut insert_at = chart.len();
    visit_fields(chart, &mut budget, |field| {
        if field.number == SERIES_NON_STYLES_FIELD {
            existing = Some(field);
        } else if field.number > SERIES_NON_STYLES_FIELD && insert_at == chart.len() {
            insert_at = field.start;
        }
        Ok(())
    })?;
    let next_chart = match existing {
        Some(field) if field.payload(chart) == sparse.as_slice() => {
            return Ok(Rewrite {
                bytes: Cow::Borrowed(chart_drawable),
            });
        },
        Some(field) => replace_payload(chart, &field, &sparse, limits)?,
        None => {
            let mut output = allocate(chart.len() + sparse.len() + 7, limits)?;
            output.extend_from_slice(&chart[..insert_at]);
            push_varint(&mut output, (u64::from(SERIES_NON_STYLES_FIELD) << 3) | 2);
            push_varint(&mut output, sparse.len() as u64);
            output.extend_from_slice(&sparse);
            output.extend_from_slice(&chart[insert_at..]);
            output
        },
    };
    let output = replace_payload(chart_drawable, &chart_field, &next_chart, limits)?;
    Ok(Rewrite {
        bytes: Cow::Owned(output),
    })
}

/// `TSCH.ChartArchive.reference_lines` (extension field 10005).
pub const REFERENCE_LINES_FIELD: u32 = 10_005;

/// The payload of a chart's `reference_lines` extension (a
/// `TSCH.ChartReferenceLinesArchive`), if stored.
///
/// # Errors
///
/// [`Error::Malformed`] for a chart drawable without a chart extension;
/// framing, duplicate-field, wire-type and limit errors.
pub fn reference_lines(chart_drawable: &[u8], limits: Limits) -> Result<Option<&[u8]>, Error> {
    check_bytes(chart_drawable.len(), limits)?;
    let mut budget = Budget::new(limits);
    let chart_field = chart_extension(chart_drawable, &mut budget)?;
    let chart = chart_field.payload(chart_drawable);
    let Some(field) = singular(chart, REFERENCE_LINES_FIELD, &mut budget)? else {
        return Ok(None);
    };
    require_wire_type(&field, WIRE_LENGTH_DELIMITED)?;
    Ok(Some(field.payload(chart)))
}

/// Set (`Some(payload)`) or remove (`None`) a chart's `reference_lines`
/// extension: a stored field is replaced in place; an absent one is
/// inserted before the first chart field numbered above 10005. Every other
/// byte is kept.
///
/// # Errors
///
/// The [`reference_lines`] errors; [`Error::Limit`]/[`Error::Allocation`]
/// for the candidate.
pub fn rewrite_reference_lines<'source>(
    chart_drawable: &'source [u8],
    payload: Option<&[u8]>,
    limits: Limits,
) -> Result<Rewrite<'source>, Error> {
    reference_lines(chart_drawable, limits)?;
    let mut budget = Budget::new(limits);
    let chart_field = chart_extension(chart_drawable, &mut budget)?;
    let chart = chart_field.payload(chart_drawable);
    let mut existing = None;
    let mut insert_at = chart.len();
    visit_fields(chart, &mut budget, |field| {
        if field.number == REFERENCE_LINES_FIELD {
            existing = Some(field);
        } else if field.number > REFERENCE_LINES_FIELD && insert_at == chart.len() {
            insert_at = field.start;
        }
        Ok(())
    })?;
    let next_chart = match (existing, payload) {
        (None, None) => None,
        (Some(field), None) => Some(remove_field(chart, &field, limits)?),
        (Some(field), Some(payload)) if field.payload(chart) == payload => None,
        (Some(field), Some(payload)) => Some(replace_payload(chart, &field, payload, limits)?),
        (None, Some(payload)) => {
            let mut output = allocate(chart.len() + payload.len() + 8, limits)?;
            output.extend_from_slice(&chart[..insert_at]);
            push_varint(&mut output, (u64::from(REFERENCE_LINES_FIELD) << 3) | 2);
            push_varint(&mut output, payload.len() as u64);
            output.extend_from_slice(payload);
            output.extend_from_slice(&chart[insert_at..]);
            Some(output)
        },
    };
    let Some(next_chart) = next_chart else {
        return Ok(Rewrite {
            bytes: Cow::Borrowed(chart_drawable),
        });
    };
    let output = replace_payload(chart_drawable, &chart_field, &next_chart, limits)?;
    Ok(Rewrite {
        bytes: Cow::Owned(output),
    })
}

/// The identifier of one `TSP.Reference` field (`None` for 0).
fn reference(data: &[u8], field: &Field, budget: &mut Budget) -> Result<Option<u64>, Error> {
    require_wire_type(field, WIRE_LENGTH_DELIMITED)?;
    let message = field.payload(data);
    let identifier = singular(message, REFERENCE_IDENTIFIER_FIELD, budget)?;
    let Some(identifier) = identifier else {
        return Ok(None);
    };
    require_wire_type(&identifier, WIRE_VARINT)?;
    Ok((identifier.varint != 0).then_some(identifier.varint))
}

fn sparse_references(data: &[u8], budget: &mut Budget) -> Result<SparseReferences, Error> {
    let mut sparse = SparseReferences::default();
    let mut count_seen = false;
    let mut entries = Vec::new();
    visit_fields(data, budget, |field| {
        match field.number {
            SPARSE_COUNT_FIELD => {
                require_wire_type(&field, WIRE_VARINT)?;
                if std::mem::replace(&mut count_seen, true) {
                    return Err(Error::DuplicateField {
                        field: field.number,
                    });
                }
                sparse.count = field.varint as u32;
            },
            SPARSE_ENTRIES_FIELD => {
                require_wire_type(&field, WIRE_LENGTH_DELIMITED)?;
                entries.push(field);
            },
            _ => {},
        }
        Ok(())
    })?;
    for entry in entries {
        let entry = entry.payload(data);
        let mut index = None;
        let mut reference_field = None;
        visit_fields(entry, budget, |field| {
            match field.number {
                SPARSE_ENTRY_INDEX_FIELD => {
                    require_wire_type(&field, WIRE_VARINT)?;
                    if index.replace(field.varint as u32).is_some() {
                        return Err(Error::DuplicateField {
                            field: field.number,
                        });
                    }
                },
                SPARSE_ENTRY_REFERENCE_FIELD => {
                    if reference_field.replace(field).is_some() {
                        return Err(Error::DuplicateField {
                            field: field.number,
                        });
                    }
                },
                _ => {},
            }
            Ok(())
        })?;
        let identifier = match reference_field {
            Some(field) => reference(entry, &field, budget)?,
            None => None,
        };
        let (Some(index), Some(identifier)) = (index, identifier) else {
            return Err(Error::Malformed {
                reason: "a sparse reference entry lacks its index or reference",
            });
        };
        sparse.entries.push((index, identifier));
    }
    Ok(sparse)
}

pub mod merge;

#[cfg(test)]
mod tests;
