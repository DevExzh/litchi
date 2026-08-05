use thiserror::Error;

use crate::Kind;

/// Result type used by the BIFF framing crate.
pub type Result<T> = std::result::Result<T, Error>;

/// A bounded resource named by a framing error.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Resource {
    /// Borrowed input bytes.
    InputBytes,
    /// Encoded output bytes.
    OutputBytes,
    /// Number of frames.
    RecordCount,
    /// Payload bytes in one frame.
    RecordBytes,
    /// A frame header.
    RecordHeader,
    /// A frame payload.
    RecordPayload,
    /// One encoded frame.
    EncodedRecord,
    /// An encoded stream.
    EncodedStream,
    /// An owned frame allocation.
    RecordFrame,
}

impl std::fmt::Display for Resource {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let name = match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::RecordCount => "record count",
            Self::RecordBytes => "record bytes",
            Self::RecordHeader => "record header",
            Self::RecordPayload => "record payload",
            Self::EncodedRecord => "encoded record",
            Self::EncodedStream => "encoded stream",
            Self::RecordFrame => "record frame",
        };
        formatter.write_str(name)
    }
}

/// A checked failure while reading, preserving, or writing BIFF frames.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// A configured bound exceeds a physical constraint of the frame format.
    #[error("invalid {resource} limit {value}; maximum is {maximum}")]
    InvalidLimit {
        /// Resource whose bound was rejected.
        resource: Resource,
        /// Requested bound.
        value: u64,
        /// Largest supported bound.
        maximum: u64,
    },

    /// Input or output crossed a configured resource bound.
    #[error("{resource} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource whose bound was crossed.
        resource: Resource,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },

    /// A value from a wider host integer domain cannot be represented as a
    /// two-byte BIFF record kind.
    #[error("BIFF record kind value {value} does not fit in an unsigned 16-bit value")]
    InvalidKind {
        /// Rejected host value.
        value: u64,
    },

    /// Fewer than four bytes remain for a frame header.
    #[error("truncated BIFF header at offset {offset}: only {available} byte(s) remain")]
    TruncatedHeader {
        /// Header offset in the input stream.
        offset: usize,
        /// Bytes available from the header offset.
        available: usize,
    },

    /// A frame declares more payload bytes than remain in the input.
    #[error(
        "truncated BIFF record {kind} at offset {offset}: declared {declared} byte(s), only {available} available"
    )]
    TruncatedPayload {
        /// Header offset in the input stream.
        offset: usize,
        /// Record kind declared by the header.
        kind: Kind,
        /// Payload length declared by the header.
        declared: usize,
        /// Payload bytes available after the header.
        available: usize,
    },

    /// An owned [`crate::Record`] was given no frame at all.
    #[error("BIFF record frame is empty")]
    EmptyRecord,

    /// An owned [`crate::Record`] was given more than one frame.
    #[error("BIFF record frame at offset {offset} contains more than one record")]
    MultipleRecords {
        /// Offset of the second complete frame.
        offset: usize,
    },

    /// Checked length arithmetic could not be represented.
    #[error("size overflow while processing {resource}")]
    SizeOverflow {
        /// Resource being sized.
        resource: Resource,
    },

    /// A fallible allocation could not reserve storage.
    #[error("could not allocate storage for {resource}")]
    Allocation {
        /// Resource being allocated.
        resource: Resource,
    },
}
