//! Inert parser for VBA projects stored according to [MS-OVBA].
//!
//! This module exposes source text and metadata for inspection. It does not
//! compile, interpret, or execute VBA.

mod compression;
mod directory;
mod project;

pub use compression::{compress_container, decompress_container};
pub use directory::{VbaDirectory, VbaModuleKind, VbaModuleMetadata};
pub use project::{VbaModule, VbaProject, VbaText};

use crate::OleError;
use std::fmt;

const DEFAULT_MAX_COMPRESSED_STREAM_BYTES: usize = 32 * 1024 * 1024;
const DEFAULT_MAX_DECOMPRESSED_STREAM_BYTES: usize = 64 * 1024 * 1024;
const DEFAULT_MAX_MODULES: usize = 4_096;
const DEFAULT_MAX_STRING_BYTES: usize = 1024 * 1024;
const DEFAULT_MAX_TOTAL_SOURCE_BYTES: usize = 128 * 1024 * 1024;

/// Resource limits applied while parsing untrusted VBA projects.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct VbaLimits {
    /// Maximum compressed size of a `dir` or module stream suffix.
    pub max_compressed_stream_bytes: usize,
    /// Maximum bytes produced from one MS-OVBA compressed container.
    pub max_decompressed_stream_bytes: usize,
    /// Maximum module count accepted from the `dir` stream.
    pub max_modules: usize,
    /// Maximum byte length of one string record.
    pub max_string_bytes: usize,
    /// Maximum aggregate decompressed module-source size.
    pub max_total_source_bytes: usize,
}

impl Default for VbaLimits {
    fn default() -> Self {
        Self {
            max_compressed_stream_bytes: DEFAULT_MAX_COMPRESSED_STREAM_BYTES,
            max_decompressed_stream_bytes: DEFAULT_MAX_DECOMPRESSED_STREAM_BYTES,
            max_modules: DEFAULT_MAX_MODULES,
            max_string_bytes: DEFAULT_MAX_STRING_BYTES,
            max_total_source_bytes: DEFAULT_MAX_TOTAL_SOURCE_BYTES,
        }
    }
}

/// Error returned for malformed or over-limit MS-OVBA data.
#[derive(Debug)]
pub enum VbaError {
    /// The containing CFB file could not be read.
    Ole(OleError),
    /// A required MS-OVBA structure is malformed.
    InvalidData(String),
    /// An explicit parser resource limit was exceeded.
    LimitExceeded {
        /// Name of the exceeded limit.
        limit: &'static str,
        /// Observed or requested byte/item count.
        actual: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// The project declares a code page unsupported by the encoding layer.
    UnsupportedCodePage(u16),
}

impl fmt::Display for VbaError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Ole(error) => write!(f, "{error}"),
            Self::InvalidData(message) => write!(f, "invalid MS-OVBA data: {message}"),
            Self::LimitExceeded {
                limit,
                actual,
                maximum,
            } => write!(f, "{limit} limit exceeded: {actual} > {maximum}"),
            Self::UnsupportedCodePage(code_page) => {
                write!(f, "unsupported VBA project code page {code_page}")
            },
        }
    }
}

impl std::error::Error for VbaError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Ole(error) => Some(error),
            _ => None,
        }
    }
}

impl From<OleError> for VbaError {
    fn from(error: OleError) -> Self {
        Self::Ole(error)
    }
}

pub(crate) fn invalid(message: impl Into<String>) -> VbaError {
    VbaError::InvalidData(message.into())
}

pub(crate) fn check_limit(
    limit: &'static str,
    actual: usize,
    maximum: usize,
) -> Result<(), VbaError> {
    if actual > maximum {
        return Err(VbaError::LimitExceeded {
            limit,
            actual,
            maximum,
        });
    }
    Ok(())
}
