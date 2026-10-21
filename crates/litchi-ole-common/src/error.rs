//! Typed failures at the bounded object-capture boundary.

use std::{error, fmt, io};

use litchi_cfb::OleError;

/// Finite resource dimensions enforced by object capture and editing.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum LimitKind {
    /// Selected object count.
    Objects,
    /// Storage path depth.
    StorageDepth,
    /// Streams in one selected object.
    ObjectStreams,
    /// Streams retained in the complete package.
    PackageStreams,
    /// Bytes in one stream.
    StreamBytes,
    /// Bytes in one rendered object.
    ObjectBytes,
    /// Aggregate retained stream bytes.
    TotalBytes,
    /// Descendant storages in one selected object.
    ObjectStorages,
    /// Storages retained in the complete package.
    PackageStorages,
    /// Directory entries visited under an object capture budget.
    DirectoryEntries,
    /// Stream removals in one operation.
    StreamRemovals,
    /// Components in a stream selector.
    SelectorDepth,
}

/// Object errors preserve low-level CFB failures and identify object budgets.
#[derive(Debug)]
pub enum Error {
    /// A low-level CFB or raw metadata failure.
    Cfb(OleError),
    /// A finite object resource budget was exceeded.
    Limit {
        /// Resource dimension.
        kind: LimitKind,
        /// Observed resource usage.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// An object limit is outside its supported nonzero range.
    InvalidLimit {
        /// Resource dimension.
        kind: LimitKind,
        /// Requested value.
        value: u64,
        /// Largest supported value.
        maximum: u64,
    },
    /// The recipient does not match the patch's recorded base.
    PatchConflict,
}

impl From<OleError> for Error {
    #[cold]
    #[inline(never)]
    fn from(error: OleError) -> Self {
        match error {
            OleError::LimitExceeded {
                resource,
                observed,
                maximum,
            } => Self::Limit {
                kind: limit_kind(resource),
                observed,
                maximum,
            },
            OleError::InvalidLimit {
                resource,
                value,
                maximum,
            } => Self::InvalidLimit {
                kind: limit_kind(resource),
                value,
                maximum,
            },
            error => Self::Cfb(error),
        }
    }
}

/// Maps the fixed low-level resource labels used by the shared raw helpers to
/// the common dimensions exposed by their public errors.  Unknown CFB labels
/// remain stream-sized input because the substrate only supplies that
/// physical dimension to this boundary.
fn limit_kind(resource: &str) -> LimitKind {
    if resource.contains("directory")
        || resource.contains("entries")
        || resource.contains("TOC")
        || resource.contains("toc")
    {
        LimitKind::DirectoryEntries
    } else if resource.contains("depth") {
        LimitKind::StorageDepth
    } else {
        LimitKind::StreamBytes
    }
}

impl From<io::Error> for Error {
    fn from(error: io::Error) -> Self {
        Self::Cfb(OleError::Io(error))
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Cfb(error) => error.fmt(formatter),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => {
                write!(
                    formatter,
                    "OLE {kind:?} budget exceeded: {observed} > {maximum}"
                )
            },
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => {
                write!(
                    formatter,
                    "invalid OLE {kind:?} limit: {value}; supported range 1..={maximum}"
                )
            },
            Self::PatchConflict => formatter.write_str("OLE patch base does not match recipient"),
        }
    }
}

impl error::Error for Error {
    fn source(&self) -> Option<&(dyn error::Error + 'static)> {
        match self {
            Self::Cfb(error) => Some(error),
            _ => None,
        }
    }
}

impl From<Error> for OleError {
    fn from(error: Error) -> Self {
        match error {
            Error::Cfb(error) => error,
            Error::Limit {
                kind,
                observed,
                maximum,
            } => OleError::LimitExceeded {
                resource: kind.resource_name(),
                observed,
                maximum,
            },
            Error::InvalidLimit {
                kind,
                value,
                maximum,
            } => OleError::InvalidLimit {
                resource: kind.resource_name(),
                value,
                maximum,
            },
            Error::PatchConflict => OleError::InvalidFormat(
                "OLE patch base does not match recipient".to_string(),
            ),
        }
    }
}

impl LimitKind {
    fn resource_name(self) -> &'static str {
        match self {
            Self::Objects => "objects",
            Self::StorageDepth => "storage depth",
            Self::ObjectStreams => "object streams",
            Self::PackageStreams => "package streams",
            Self::StreamBytes => "stream bytes",
            Self::ObjectBytes => "object bytes",
            Self::TotalBytes => "total bytes",
            Self::ObjectStorages => "object storages",
            Self::PackageStorages => "package storages",
            Self::DirectoryEntries => "directory entries",
            Self::StreamRemovals => "stream removals",
            Self::SelectorDepth => "selector depth",
        }
    }

    fn resource(self) -> litchi_core::Resource {
        match self {
            Self::StorageDepth | Self::SelectorDepth => litchi_core::Resource::Depth,
            Self::StreamBytes | Self::ObjectBytes | Self::TotalBytes => {
                litchi_core::Resource::InputBytes
            },
            Self::Objects
            | Self::ObjectStreams
            | Self::PackageStreams
            | Self::ObjectStorages
            | Self::PackageStorages
            | Self::DirectoryEntries
            | Self::StreamRemovals => litchi_core::Resource::Objects,
        }
    }
}

impl From<Error> for litchi_core::Error {
    fn from(error: Error) -> Self {
        match error {
            Error::Cfb(error @ OleError::SourceChanged { .. }) => Self::Io(io::Error::other(error)),
            Error::Cfb(OleError::InvalidLimit { value, maximum, .. }) => Self::InvalidLimit {
                kind: litchi_core::Resource::InputBytes,
                value,
                maximum,
                scope: "OLE container",
            },
            Error::Cfb(error) => error.into(),
            Error::Limit {
                kind,
                observed,
                maximum,
            } => Self::ResourceLimit(litchi_core::ResourceLimit {
                resource: kind.resource(),
                observed,
                limit: maximum,
                scope: format!("OLE {kind:?}").into(),
            }),
            Error::InvalidLimit {
                kind,
                value,
                maximum,
            } => Self::InvalidLimit {
                kind: kind.resource(),
                value,
                maximum,
                scope: "OLE object",
            },
            Error::PatchConflict => Self::PatchConflict,
        }
    }
}
