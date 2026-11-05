//! Structured failures produced while validating `OfficeArt` data.

use crate::record::RecordKind;

/// The bounded traversal resource that was exhausted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LimitKind {
    /// Maximum body bytes accepted by one raw record.
    RecordBytes,
    /// Maximum nested-container depth.
    Depth,
    /// Maximum number of visited records.
    Records,
}

/// A checked `OfficeArt` parsing or validation failure.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// Fewer than eight bytes were available for a record header.
    TruncatedHeader {
        /// Requested header offset.
        offset: usize,
        /// Bytes available from the requested offset.
        available: usize,
    },
    /// `recLen` extends past the supplied byte slice.
    TruncatedPayload {
        /// Offset of the record header.
        offset: usize,
        /// Payload length declared by `recLen`.
        declared: u32,
        /// Payload bytes actually available.
        available: usize,
    },
    /// Checked address or length arithmetic overflowed.
    ArithmeticOverflow {
        /// Operation that could not be represented.
        context: &'static str,
    },
    /// A record passed to a container API cannot contain children.
    NotContainer {
        /// Semantic record kind.
        kind: RecordKind,
        /// Exact kind value read from the wire.
        raw_kind: u16,
    },
    /// An `OfficeArt` property table is structurally invalid.
    MalformedProperties {
        /// Concise validation failure description.
        reason: &'static str,
    },
    /// A custom `OfficeArt` geometry property violates `[MS-ODRAW]`.
    MalformedGeometry {
        /// Concise validation failure description.
        reason: &'static str,
    },
    /// A shape container violates an `OfficeArt` structural invariant.
    MalformedShape {
        /// Concise validation failure description.
        reason: &'static str,
    },
    /// An `OfficeArt` BLIP, FBSE, or BLIP-store invariant is invalid.
    MalformedImage {
        /// Concise validation failure description.
        reason: &'static str,
    },
    /// A record is not an `OfficeArt` image record.
    NotImageRecord {
        /// Exact record type read from the wire.
        raw_kind: u16,
    },
    /// A caller supplied an invalid one-based BLIP-store identifier.
    InvalidBlipId {
        /// Rejected numeric identifier.
        value: u32,
    },
    /// An image record exceeded an explicit resource ceiling.
    ImageLimitExceeded {
        /// Resource whose ceiling was reached.
        limit: crate::image::LimitKind,
        /// Observed resource size or count.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A delay-loaded FBSE was resolved without its associated delay store.
    MissingDelayStore,
    /// An FBSE delay offset is outside the supplied delay store.
    DelayOffsetOutOfBounds {
        /// Requested byte offset.
        offset: u32,
        /// Delay-store byte length.
        available: usize,
    },
    /// An image length field does not match the bytes it describes.
    ImageSizeMismatch {
        /// Field being validated.
        field: &'static str,
        /// Length declared on the wire.
        declared: u64,
        /// Actual byte length.
        actual: usize,
    },
    /// A single-root parser was given additional top-level bytes.
    TrailingData {
        /// Offset at which the unexpected top-level data begins.
        offset: usize,
    },
    /// A caller-supplied traversal limit exceeds the implementation's safe bound.
    InvalidLimit {
        /// Resource whose requested ceiling is unsafe.
        limit: LimitKind,
        /// Requested ceiling.
        value: u32,
        /// Largest accepted ceiling.
        maximum: u32,
    },
    /// A bounded recursive traversal reached its configured ceiling.
    LimitExceeded {
        /// Resource whose limit was reached.
        limit: LimitKind,
        /// Observed resource size or count.
        observed: u64,
        /// Configured maximum.
        maximum: u32,
    },
}

impl core::fmt::Display for Error {
    fn fmt(&self, formatter: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        match self {
            Self::TruncatedHeader { offset, available } => write!(
                formatter,
                "OfficeArt record at offset {offset} needs an 8-byte header; {available} bytes remain"
            ),
            Self::TruncatedPayload {
                offset,
                declared,
                available,
            } => write!(
                formatter,
                "OfficeArt record at offset {offset} declares {declared} payload bytes; {available} bytes remain"
            ),
            Self::ArithmeticOverflow { context } => {
                write!(formatter, "OfficeArt {context} cannot be represented")
            },
            Self::NotContainer { kind, raw_kind } => write!(
                formatter,
                "OfficeArt record {kind:?} (0x{raw_kind:04X}) is not a container"
            ),
            Self::MalformedProperties { reason } => {
                write!(formatter, "malformed OfficeArt properties: {reason}")
            },
            Self::MalformedGeometry { reason } => {
                write!(formatter, "malformed OfficeArt geometry: {reason}")
            },
            Self::MalformedShape { reason } => {
                write!(formatter, "malformed OfficeArt shape: {reason}")
            },
            Self::MalformedImage { reason } => {
                write!(formatter, "malformed OfficeArt image: {reason}")
            },
            Self::NotImageRecord { raw_kind } => {
                write!(
                    formatter,
                    "OfficeArt record 0x{raw_kind:04X} is not an image record"
                )
            },
            Self::InvalidBlipId { value } => {
                write!(
                    formatter,
                    "OfficeArt BLIP identifier {value} is outside 1..=4095"
                )
            },
            Self::ImageLimitExceeded {
                limit,
                observed,
                maximum,
            } => write!(
                formatter,
                "OfficeArt image {limit:?} observed {observed} exceeds {maximum}"
            ),
            Self::MissingDelayStore => {
                write!(
                    formatter,
                    "OfficeArt image requires an associated delay store"
                )
            },
            Self::DelayOffsetOutOfBounds { offset, available } => write!(
                formatter,
                "OfficeArt delay offset {offset} is outside a {available}-byte delay store"
            ),
            Self::ImageSizeMismatch {
                field,
                declared,
                actual,
            } => write!(
                formatter,
                "OfficeArt image {field} declares {declared} bytes; actual size is {actual}"
            ),
            Self::TrailingData { offset } => {
                write!(
                    formatter,
                    "unexpected OfficeArt top-level data at offset {offset}"
                )
            },
            Self::InvalidLimit {
                limit,
                value,
                maximum,
            } => write!(
                formatter,
                "OfficeArt {limit:?} limit {value} exceeds the safe maximum of {maximum}"
            ),
            Self::LimitExceeded {
                limit,
                observed,
                maximum,
            } => {
                write!(
                    formatter,
                    "OfficeArt {limit:?} observed {observed} exceeds {maximum}"
                )
            },
        }
    }
}

impl std::error::Error for Error {}

/// Result type for checked `OfficeArt` parsing and validation.
pub type Result<T> = core::result::Result<T, Error>;
