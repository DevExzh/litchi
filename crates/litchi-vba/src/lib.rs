//! Bounded, inert parsing and authoring of [MS-OVBA] projects.
//!
//! This crate exposes project metadata and source text for inspection and
//! deterministic serialization. It never compiles, interprets, or executes
//! Visual Basic source.

#![forbid(unsafe_code)]

pub mod build;
pub mod codec;
pub mod dir;
pub mod project;

pub use build::Payload;

use litchi_cfb::OleError;
use std::fmt;

const DEFAULT_MAX_CFB_BYTES: usize = 256 * 1024 * 1024;
const DEFAULT_MAX_COMPRESSED_STREAM_BYTES: usize = 32 * 1024 * 1024;
const DEFAULT_MAX_DECOMPRESSED_STREAM_BYTES: usize = 64 * 1024 * 1024;
const DEFAULT_MAX_MODULES: usize = 4_096;
const DEFAULT_MAX_STRING_BYTES: usize = 1024 * 1024;
const DEFAULT_MAX_TOTAL_SOURCE_BYTES: usize = 128 * 1024 * 1024;

/// Resource ceilings applied while parsing or authoring an inert project.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum standalone CFB payload size.
    pub max_cfb_bytes: usize,
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

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_cfb_bytes: DEFAULT_MAX_CFB_BYTES,
            max_compressed_stream_bytes: DEFAULT_MAX_COMPRESSED_STREAM_BYTES,
            max_decompressed_stream_bytes: DEFAULT_MAX_DECOMPRESSED_STREAM_BYTES,
            max_modules: DEFAULT_MAX_MODULES,
            max_string_bytes: DEFAULT_MAX_STRING_BYTES,
            max_total_source_bytes: DEFAULT_MAX_TOTAL_SOURCE_BYTES,
        }
    }
}

/// Error returned for malformed, unsupported, or over-limit MS-OVBA data.
#[derive(Debug)]
pub enum Error {
    /// The containing CFB file could not be read or written.
    Cfb(OleError),
    /// A required MS-OVBA structure is malformed.
    InvalidData(String),
    /// An explicit resource ceiling was exceeded.
    LimitExceeded {
        /// Name of the exceeded ceiling.
        limit: &'static str,
        /// Observed or requested byte/item count.
        actual: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// The project declares a code page unsupported by the encoding layer.
    UnsupportedCodePage(u16),
}

impl Error {
    /// Create a malformed-data error with contextual detail.
    pub fn invalid(message: impl Into<String>) -> Self {
        Self::InvalidData(message.into())
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Cfb(error) => write!(formatter, "{error}"),
            Self::InvalidData(message) => write!(formatter, "invalid MS-OVBA data: {message}"),
            Self::LimitExceeded {
                limit,
                actual,
                maximum,
            } => write!(formatter, "{limit} limit exceeded: {actual} > {maximum}"),
            Self::UnsupportedCodePage(code_page) => {
                write!(formatter, "unsupported VBA project code page {code_page}")
            },
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Cfb(error) => Some(error),
            Self::InvalidData(_) | Self::LimitExceeded { .. } | Self::UnsupportedCodePage(_) => {
                None
            },
        }
    }
}

impl From<OleError> for Error {
    fn from(error: OleError) -> Self {
        Self::Cfb(error)
    }
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::invalid(message)
}

pub(crate) fn check_limit(limit: &'static str, actual: usize, maximum: usize) -> Result<(), Error> {
    if actual > maximum {
        return Err(Error::LimitExceeded {
            limit,
            actual,
            maximum,
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn owned_outputs_are_send_and_sync() {
        assert_send_sync::<Payload>();
        assert_send_sync::<project::Project>();
    }
}
