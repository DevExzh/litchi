use std::fmt;

/// A checked half-open byte range within one fragment.
///
/// The end offset is stored rather than recomputed so every accessor remains
/// overflow-free after construction. Empty spans are valid because an adapter
/// may need to index a present, zero-length payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct ByteSpan {
    start: u64,
    end: u64,
}

impl ByteSpan {
    /// Construct a span from an offset and a byte length.
    ///
    /// # Errors
    ///
    /// Returns [`ByteSpanError::Overflow`] when the exclusive end offset
    /// cannot be represented by `u64`.
    pub const fn new(start: u64, length: u64) -> Result<Self, ByteSpanError> {
        match start.checked_add(length) {
            Some(end) => Ok(Self { start, end }),
            None => Err(ByteSpanError::Overflow { start, length }),
        }
    }

    /// Construct a span from checked half-open endpoints.
    ///
    /// # Errors
    ///
    /// Returns [`ByteSpanError::Reversed`] when `end` precedes `start`.
    pub const fn from_endpoints(start: u64, end: u64) -> Result<Self, ByteSpanError> {
        if end < start {
            Err(ByteSpanError::Reversed { start, end })
        } else {
            Ok(Self { start, end })
        }
    }

    /// Return the first byte offset in the span.
    #[must_use]
    pub const fn start(self) -> u64 {
        self.start
    }

    /// Return the exclusive byte offset after the span.
    #[must_use]
    pub const fn end(self) -> u64 {
        self.end
    }

    /// Return the number of bytes in the span.
    #[must_use]
    pub const fn length(self) -> u64 {
        self.end - self.start
    }

    /// Return whether the span contains no bytes.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.start == self.end
    }
}

/// Failure while constructing a checked byte span.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ByteSpanError {
    /// The end offset would overflow `u64`.
    Overflow { start: u64, length: u64 },
    /// The supplied end precedes the supplied start.
    Reversed { start: u64, end: u64 },
}

impl fmt::Display for ByteSpanError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Overflow { start, length } => write!(
                formatter,
                "byte span start {start} plus length {length} overflows"
            ),
            Self::Reversed { start, end } => {
                write!(formatter, "byte span end {end} precedes start {start}")
            },
        }
    }
}

impl std::error::Error for ByteSpanError {}
