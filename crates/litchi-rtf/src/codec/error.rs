//! Error types for RTF parsing.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "items stay grouped by RTF feature area rather than by item kind"
)]
use std::fmt;
use std::hash::{BuildHasher, Hash};

use super::limits::SourceLimitKind;

/// Result type for RTF operations.
pub type Result<T> = std::result::Result<T, Error>;

/// Reserve additional vector elements before an atomic model mutation.
pub(crate) fn try_reserve_additional<T>(
    values: &mut Vec<T>,
    additional: usize,
    resource: &'static str,
) -> Result<()> {
    let requested = values
        .len()
        .saturating_add(additional)
        .saturating_mul(size_of::<T>());
    values
        .try_reserve(additional)
        .map_err(|_err| Error::AllocationFailed {
            resource,
            requested,
        })
}

/// Reserve one additional vector element without exposing an allocation panic
/// as a recoverable model mutation.
pub(crate) fn try_reserve_one<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    try_reserve_additional(values, 1, resource)
}

/// Reserve hash-set entries before validation inserts untrusted collection
/// members.
pub(crate) fn try_reserve_set<T, S>(
    values: &mut std::collections::HashSet<T, S>,
    additional: usize,
    resource: &'static str,
) -> Result<()>
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
        .map_err(|_err| Error::AllocationFailed {
            resource,
            requested,
        })
}

/// RTF parsing errors.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum Error {
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
    /// The document's `\ansicpg` names a code page with no exact codec,
    /// such as the Mac Japanese (10001) and Mac Cyrillic (10007) pages.
    /// The document is well formed; its text cannot be decoded exactly.
    UnsupportedCodePage(u32),
    /// Well-formed content that the retained model cannot represent faithfully.
    UnsupportedContent(&'static str),
    /// A source-limit value was invalid at profile construction.
    InvalidLimit {
        /// Stable physical ingress dimension.
        kind: SourceLimitKind,
        /// Caller-provided value.
        value: usize,
        /// Maximum addressable value for the dimension.
        maximum: usize,
    },
    /// A configured physical ingress budget was exceeded.
    Limit {
        /// Stable physical ingress dimension.
        kind: SourceLimitKind,
        /// Value observed by the parser or transport decoder.
        observed: usize,
        /// Configured maximum value.
        maximum: usize,
    },
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

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::LexerError(msg) => write!(f, "RTF Lexer Error: {msg}"),
            Error::ParserError(msg) => write!(f, "RTF Parser Error: {msg}"),
            Error::InvalidStructure(msg) => write!(f, "Invalid RTF structure: {msg}"),
            Error::InvalidUnicode(msg) => write!(f, "Invalid unicode: {msg}"),
            Error::UnexpectedEof => write!(f, "Unexpected end of input"),
            Error::InvalidControlWord(msg) => write!(f, "Invalid control word: {msg}"),
            Error::MalformedDocument(msg) => write!(f, "Malformed RTF document: {msg}"),
            Error::UnsupportedCodePage(page) => {
                write!(f, "unsupported RTF ANSI code page {page}")
            },
            Error::UnsupportedContent(content) => write!(f, "unsupported RTF content: {content}"),
            Error::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(
                f,
                "invalid RTF source limit for {kind}: value {value}, maximum {maximum}"
            ),
            Error::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                f,
                "RTF source limit exceeded for {kind}: observed {observed}, maximum {maximum}"
            ),
            Error::LimitExceeded {
                resource,
                observed,
                limit,
            } => write!(
                f,
                "RTF resource limit exceeded for {resource}: observed {observed}, limit {limit}"
            ),
            Error::AllocationFailed {
                resource,
                requested,
            } => write!(
                f,
                "RTF allocation failed for {resource}: requested {requested} bytes"
            ),
        }
    }
}

impl std::error::Error for Error {}

impl From<std::str::Utf8Error> for Error {
    fn from(err: std::str::Utf8Error) -> Self {
        Error::InvalidUnicode(err.to_string())
    }
}

impl From<std::num::ParseIntError> for Error {
    fn from(err: std::num::ParseIntError) -> Self {
        Error::ParserError(format!("Integer parsing error: {err}"))
    }
}
