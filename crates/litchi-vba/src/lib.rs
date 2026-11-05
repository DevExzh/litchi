//! Bounded, inert parsing and authoring of [MS-OVBA] projects.
//!
//! This crate exposes project metadata and source text for inspection and
//! deterministic serialization. It never compiles, interprets, or executes
//! Visual Basic source.

#![forbid(unsafe_code)]

pub mod build;
pub mod codec;
pub mod dir;
mod limit;
pub mod project;

pub use build::Payload;
pub use limit::LimitKind;

use litchi_cfb::OleError;
use std::fmt;

const DEFAULT_MAX_CFB_BYTES: usize = 256 * 1024 * 1024;
const DEFAULT_MAX_COMPRESSED_STREAM_BYTES: usize = 32 * 1024 * 1024;
const DEFAULT_MAX_DECOMPRESSED_STREAM_BYTES: usize = 64 * 1024 * 1024;
const DEFAULT_MAX_MODULES: usize = 4_096;
const DEFAULT_MAX_REFERENCES: usize = 4_096;
const DEFAULT_MAX_LICENSES: usize = 4_096;
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
    /// Maximum external-reference count accepted from the `dir` stream.
    pub max_references: usize,
    /// Maximum ActiveX license-record count accepted from `PROJECTlk`.
    pub max_licenses: usize,
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
            max_references: DEFAULT_MAX_REFERENCES,
            max_licenses: DEFAULT_MAX_LICENSES,
            max_string_bytes: DEFAULT_MAX_STRING_BYTES,
            max_total_source_bytes: DEFAULT_MAX_TOTAL_SOURCE_BYTES,
        }
    }
}

/// Error returned for malformed, unsupported, or over-limit MS-OVBA data.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The containing CFB file could not be read or written.
    Cfb(OleError),
    /// A required MS-OVBA structure is malformed.
    InvalidData(String),
    /// An explicit resource ceiling was exceeded (record 0360).
    Limit {
        /// The exceeded ceiling.
        kind: LimitKind,
        /// Observed or requested byte/item count.
        observed: usize,
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
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(formatter, "{kind} limit exceeded: {observed} > {maximum}"),
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
            Self::InvalidData(_) | Self::Limit { .. } | Self::UnsupportedCodePage(_) => None,
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

pub(crate) fn check_limit(kind: LimitKind, actual: usize, maximum: usize) -> Result<(), Error> {
    if actual > maximum {
        return Err(Error::Limit {
            kind,
            observed: actual,
            maximum,
        });
    }
    Ok(())
}

/// Derive the physical CFB reader policy from the selected VBA limits.
///
/// The MS-OVBA limits own the standalone CFB byte budget.  Reusing the CFB
/// reader's defaults here would silently replace that caller-selected budget
/// at the nested allocation boundary, so the directory and allocation-table
/// ceilings are constrained by the same finite input ceiling.
pub(crate) fn cfb_limits(limits: &Limits) -> Result<litchi_cfb::OleFileLimits, Error> {
    let maximum = u64::try_from(limits.max_cfb_bytes)
        .map_err(|_| invalid("VBA CFB byte limit does not fit in u64"))?;
    let directory_maximum = maximum.min(litchi_cfb::OleFileLimits::MAX_DIRECTORY_BYTES);
    let allocation_maximum = maximum.min(litchi_cfb::OleFileLimits::MAX_ALLOCATION_TABLE_BYTES);
    Ok(litchi_cfb::OleFileLimits::new(maximum)?
        .with_max_directory_bytes(directory_maximum)?
        .with_max_allocation_table_bytes(allocation_maximum)?)
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

    #[test]
    fn nested_cfb_limits_follow_the_selected_vba_input_budget() {
        let limits = Limits {
            max_cfb_bytes: 8 * 1024 * 1024,
            ..Limits::default()
        };
        let cfb = cfb_limits(&limits).expect("finite CFB policy");
        assert_eq!(cfb.max_input_bytes(), limits.max_cfb_bytes as u64);
        assert_eq!(cfb.max_directory_bytes(), limits.max_cfb_bytes as u64);
        assert_eq!(
            cfb.max_allocation_table_bytes(),
            limits.max_cfb_bytes as u64
        );
    }
}
