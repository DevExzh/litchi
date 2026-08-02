//! BIFF12 wire errors.

use std::io;

use thiserror::Error;

/// Wire location being decoded when input ended.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Stage {
    /// The record-kind field.
    Kind,
    /// The record-length field.
    Length,
    /// The declared record payload.
    Payload,
    /// A scalar or string inside a record payload.
    Value,
}

/// Failure while validating or writing BIFF12 wire data.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The input ended after a record had started.
    #[error("truncated {stage:?} at byte {offset}: needed {needed} bytes, found {available}")]
    Truncated {
        /// Wire component that was incomplete.
        stage: Stage,
        /// Absolute or cursor-relative byte offset.
        offset: usize,
        /// Required byte count from this point.
        needed: usize,
        /// Available byte count from this point.
        available: usize,
    },
    /// A record kind used a third byte, an overlong two-byte form, or exceeded
    /// the 14-bit BIFF12 domain.
    #[error("invalid BIFF12 record kind encoding at byte {offset}")]
    InvalidKind {
        /// Offset of the kind field.
        offset: usize,
    },
    /// A caller attempted to construct a record kind outside the wire domain.
    #[error("record kind {value:#06x} is outside the BIFF12 14-bit domain")]
    KindOutOfRange {
        /// Rejected numeric kind.
        value: u16,
    },
    /// A declared or emitted payload exceeded the configured budget.
    #[error("record payload length {length} exceeds limit {limit} at byte {offset}")]
    PayloadLimit {
        /// Declared or requested payload size.
        length: usize,
        /// Configured maximum.
        limit: usize,
        /// Record offset, or zero for writer-side validation.
        offset: usize,
    },
    /// A decoded or emitted string exceeded the configured UTF-16-unit budget.
    #[error("UTF-16 string length {units} exceeds limit {limit} at byte {offset}")]
    StringLimit {
        /// Declared or requested UTF-16 code-unit count.
        units: usize,
        /// Configured maximum.
        limit: usize,
        /// Cursor offset of the length field, or zero when writing.
        offset: usize,
    },
    /// A UTF-16LE string contained an unpaired surrogate.
    #[error("invalid UTF-16LE string at byte {offset}")]
    InvalidUtf16 {
        /// Cursor-relative byte offset of the string payload.
        offset: usize,
    },
    /// A 32-bit Boolean contained a value other than zero or one.
    #[error("invalid 32-bit Boolean {value} at byte {offset}")]
    InvalidBool {
        /// Rejected numeric Boolean.
        value: u32,
        /// Cursor-relative offset of the value.
        offset: usize,
    },
    /// A number cannot be represented exactly by any RK encoding.
    #[error("floating-point value {bits:#018x} cannot be represented exactly as RK")]
    UnrepresentableRk {
        /// IEEE-754 bits of the rejected value.
        bits: u64,
    },
    /// A strict payload parser left bytes unconsumed.
    #[error("{context}: {remaining} trailing payload bytes at byte {offset}")]
    Trailing {
        /// Diagnostic context supplied when the payload cursor was created.
        context: &'static str,
        /// Cursor-relative offset of the trailing bytes.
        offset: usize,
        /// Number of bytes not consumed.
        remaining: usize,
    },
    /// An encoded value cannot fit its required fixed-width length field.
    #[error("{what} length {length} cannot be represented on the BIFF12 wire")]
    LengthOverflow {
        /// Value whose length overflowed.
        what: &'static str,
        /// Observed length.
        length: usize,
    },
    /// The destination rejected output.
    #[error("I/O error: {0}")]
    Io(#[from] io::Error),
}

/// Result returned by BIFF12 raw operations.
pub type Result<T> = std::result::Result<T, Error>;
