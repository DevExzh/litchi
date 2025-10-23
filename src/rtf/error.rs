//! Error types for RTF parsing.

use std::fmt;

/// Result type for RTF operations.
pub type RtfResult<T> = Result<T, RtfError>;

/// RTF parsing errors.
#[derive(Debug, Clone)]
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
}

impl fmt::Display for RtfError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            RtfError::LexerError(msg) => write!(f, "RTF Lexer Error: {}", msg),
            RtfError::ParserError(msg) => write!(f, "RTF Parser Error: {}", msg),
            RtfError::InvalidStructure(msg) => write!(f, "Invalid RTF structure: {}", msg),
            RtfError::InvalidUnicode(msg) => write!(f, "Invalid unicode: {}", msg),
            RtfError::UnexpectedEof => write!(f, "Unexpected end of input"),
            RtfError::InvalidControlWord(msg) => write!(f, "Invalid control word: {}", msg),
            RtfError::MalformedDocument(msg) => write!(f, "Malformed RTF document: {}", msg),
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
        RtfError::ParserError(format!("Integer parsing error: {}", err))
    }
}
