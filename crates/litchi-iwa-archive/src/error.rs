use std::fmt;

use litchi_core::SourceVersion;

/// A resource category enforced before an ingress allocation or decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LimitKind {
    /// Complete bytes read from one path or in-memory package.
    InputBytes,
    /// Complete bytes materialized by one edited-package reassembly.
    OutputBytes,
    /// Non-directory ZIP members in one central directory.
    Entries,
    /// Declared uncompressed size of one ZIP member.
    EntryBytes,
    /// Aggregate declared uncompressed ZIP size.
    TotalBytes,
    /// Decompressed bytes in one IWA component.
    IwaStreamBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "ZIP entries",
            Self::EntryBytes => "ZIP entry bytes",
            Self::TotalBytes => "ZIP total bytes",
            Self::IwaStreamBytes => "IWA stream bytes",
        })
    }
}

/// Errors raised at the physical iWork ingress boundary.
#[derive(Debug, thiserror::Error)]
pub enum Error {
    /// A filesystem operation failed.
    #[error("iWork archive I/O error: {0}")]
    Io(#[from] std::io::Error),

    /// The ZIP implementation rejected the container.
    #[error("iWork ZIP archive error: {message}")]
    Zip { message: String },

    /// The neutral IWA framing layer rejected a decompressed component.
    #[error("iWork IWA archive error: {0}")]
    Iwa(#[from] litchi_iwa_core::Error),

    /// A caller supplied an invalid or contradictory resource profile.
    #[error("invalid iWork archive limits: {0}")]
    InvalidLimits(String),

    /// An input exceeded a checked physical resource ceiling.
    #[error("iWork archive {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category that was exceeded.
        kind: LimitKind,
        /// Observed input size/count.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },

    /// A bounded collection could not reserve its next entry.
    #[error("iWork archive allocation failed for {resource}: {amount}")]
    Allocation {
        /// Collection or buffer being allocated.
        resource: &'static str,
        /// Number of entries or bytes requested.
        amount: usize,
    },

    /// The package uses an encrypted iWork container marker.
    #[error("password-protected iWork documents are not supported")]
    Encrypted,

    /// The positional source changed while the package was being snapshotted.
    #[error(
        "iWork archive source changed during read (expected {expected:?}, observed {observed:?})"
    )]
    SourceChanged {
        /// Source identity and revision captured before reading.
        expected: SourceVersion,
        /// Source identity and revision observed after reading.
        observed: SourceVersion,
    },

    /// The bounded reassembler cannot safely preserve this physical ZIP
    /// layout or edit this member without changing semantics.
    #[error("iWork archive reassembly is unsupported: {0}")]
    Reassembly(String),

    /// The input is a ZIP container but not a valid iWork bundle shape.
    #[error("invalid iWork bundle: {0}")]
    InvalidBundle(String),
}

impl From<soapberry_zip::Error> for Error {
    fn from(error: soapberry_zip::Error) -> Self {
        Self::Zip {
            message: error.to_string(),
        }
    }
}

/// Result alias for physical iWork ingress operations.
pub type Result<T> = std::result::Result<T, Error>;
