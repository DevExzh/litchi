//! Error types for RTF parsing.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "items stay grouped by RTF feature area rather than by item kind"
)]
use std::fmt;
use std::hash::{BuildHasher, Hash};

/// Result type for RTF operations.
pub type RtfResult<T> = Result<T, RtfError>;

/// Reserve additional vector elements before an atomic model mutation.
pub(crate) fn try_reserve_additional<T>(
    values: &mut Vec<T>,
    additional: usize,
    resource: &'static str,
) -> RtfResult<()> {
    let requested = values
        .len()
        .saturating_add(additional)
        .saturating_mul(size_of::<T>());
    values
        .try_reserve(additional)
        .map_err(|_err| RtfError::AllocationFailed {
            resource,
            requested,
        })
}

/// Reserve one additional vector element without exposing an allocation panic
/// as a recoverable model mutation.
pub(crate) fn try_reserve_one<T>(values: &mut Vec<T>, resource: &'static str) -> RtfResult<()> {
    try_reserve_additional(values, 1, resource)
}

/// Reserve hash-set entries before validation inserts untrusted collection
/// members.
pub(crate) fn try_reserve_set<T, S>(
    values: &mut std::collections::HashSet<T, S>,
    additional: usize,
    resource: &'static str,
) -> RtfResult<()>
where
    T: Eq + Hash,
    S: BuildHasher,
{
    let requested = values
        .len()
        .saturating_add(additional)
        .saturating_mul(size_of::<T>());
    values
        .try_reserve(additional)
        .map_err(|_err| RtfError::AllocationFailed {
            resource,
            requested,
        })
}

/// RTF parsing errors.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum RtfError {
    /// Lexer error during tokenization
    LexerError(String),
    /// Parser error during document construction
    ParserError(String),
    /// Invalid RTF structure
    InvalidStructure(String),
    /// Invalid unicode character
    InvalidUnicode(String),
    /// Unexpected end of input
    UnexpectedEof,
    /// Invalid control word
    InvalidControlWord(String),
    /// Malformed document
    MalformedDocument(String),
    /// A finite resource budget was exceeded.
    LimitExceeded {
        /// Stable name of the exhausted resource.
        resource: &'static str,
        /// Value declared or observed by the operation.
        observed: usize,
        /// Configured maximum value.
        limit: usize,
    },
    /// A fallible allocation could not reserve the requested capacity.
    AllocationFailed {
        /// Stable name of the resource being allocated.
        resource: &'static str,
        /// Logical byte capacity requested by the operation.
        requested: usize,
    },
}

impl fmt::Display for RtfError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            RtfError::LexerError(msg) => write!(f, "RTF Lexer Error: {msg}"),
            RtfError::ParserError(msg) => write!(f, "RTF Parser Error: {msg}"),
            RtfError::InvalidStructure(msg) => write!(f, "Invalid RTF structure: {msg}"),
            RtfError::InvalidUnicode(msg) => write!(f, "Invalid unicode: {msg}"),
            RtfError::UnexpectedEof => write!(f, "Unexpected end of input"),
            RtfError::InvalidControlWord(msg) => write!(f, "Invalid control word: {msg}"),
            RtfError::MalformedDocument(msg) => write!(f, "Malformed RTF document: {msg}"),
            RtfError::LimitExceeded {
                resource,
                observed,
                limit,
            } => write!(
                f,
                "RTF resource limit exceeded for {resource}: observed {observed}, limit {limit}"
            ),
            RtfError::AllocationFailed {
                resource,
                requested,
            } => write!(
                f,
                "RTF allocation failed for {resource}: requested {requested} bytes"
            ),
        }
    }
}

impl std::error::Error for RtfError {}

impl From<std::str::Utf8Error> for RtfError {
    fn from(err: std::str::Utf8Error) -> Self {
        RtfError::InvalidUnicode(err.to_string())
    }
}

impl From<std::num::ParseIntError> for RtfError {
    fn from(err: std::num::ParseIntError) -> Self {
        RtfError::ParserError(format!("Integer parsing error: {err}"))
    }
}
