use std::fmt;

use thiserror::Error;

/// The bounded resource whose configured or observed size was invalid.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LimitKind {
    ArchiveBytes,
    Objects,
    Messages,
    MessagesPerObject,
    ObjectBytes,
    MessageBytes,
    HeaderBytes,
    HeaderFields,
    HeaderNesting,
    HeaderMemoryBytes,
    MetadataItems,
    SnappyChunkBytes,
    SnappyStreamBytes,
    SnappyCompressedChunkBytes,
    SnappyCompressedStreamBytes,
    SnappyFrames,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = match self {
            Self::ArchiveBytes => "archive byte",
            Self::Objects => "object",
            Self::Messages => "message",
            Self::MessagesPerObject => "per-object message",
            Self::ObjectBytes => "object byte",
            Self::MessageBytes => "message byte",
            Self::HeaderBytes => "header byte",
            Self::HeaderFields => "header fields",
            Self::HeaderNesting => "header nesting depth",
            Self::HeaderMemoryBytes => "decoded header memory bytes",
            Self::MetadataItems => "metadata item",
            Self::SnappyChunkBytes => "Snappy decompressed chunk bytes",
            Self::SnappyStreamBytes => "Snappy decompressed stream bytes",
            Self::SnappyCompressedChunkBytes => "Snappy compressed chunk bytes",
            Self::SnappyCompressedStreamBytes => "Snappy compressed stream bytes",
            Self::SnappyFrames => "Snappy frames",
        };
        formatter.write_str(name)
    }
}

/// Protobuf header whose private codec rejected an operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderKind {
    ArchiveInfo,
    MessageInfo,
}

impl fmt::Display for HeaderKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::ArchiveInfo => "TSP.ArchiveInfo",
            Self::MessageInfo => "TSP.MessageInfo",
        })
    }
}

/// Direction of a failed private header-codec operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderOperation {
    Decode,
    Encode,
}

impl fmt::Display for HeaderOperation {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Decode => "decode",
            Self::Encode => "encode",
        })
    }
}

/// Errors returned by the physical IWA substrate.
#[derive(Debug, Error)]
pub enum Error {
    #[error("invalid IWA archive at byte {offset}: {reason}")]
    InvalidArchive { offset: usize, reason: &'static str },

    #[error("invalid IWA limit configuration: {reason}")]
    InvalidLimits { reason: &'static str },

    #[error("IWA {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        kind: LimitKind,
        observed: usize,
        maximum: usize,
    },

    #[error("could not {operation} IWA {header} header: {reason}")]
    HeaderCodec {
        header: HeaderKind,
        operation: HeaderOperation,
        reason: String,
    },

    #[error("IWA archive reader error: {0}")]
    Io(#[from] std::io::Error),

    #[error("invalid IWA Snappy stream: {message}")]
    Snappy { message: String },

    #[error("allocation for {resource} ({requested} bytes) failed")]
    Allocation {
        resource: &'static str,
        requested: usize,
    },
}

impl Error {
    pub(crate) const fn invalid_archive(offset: usize, reason: &'static str) -> Self {
        Self::InvalidArchive { offset, reason }
    }

    pub(crate) const fn invalid_limits(reason: &'static str) -> Self {
        Self::InvalidLimits { reason }
    }

    pub(crate) const fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Self {
        Self::Limit {
            kind,
            observed,
            maximum,
        }
    }

    pub(crate) fn snappy(message: impl Into<String>) -> Self {
        Self::Snappy {
            message: message.into(),
        }
    }

    pub(crate) const fn allocation(resource: &'static str, requested: usize) -> Self {
        Self::Allocation {
            resource,
            requested,
        }
    }

    pub(crate) fn header_codec(
        header: HeaderKind,
        operation: HeaderOperation,
        reason: impl Into<String>,
    ) -> Self {
        Self::HeaderCodec {
            header,
            operation,
            reason: reason.into(),
        }
    }
}

/// Result alias for this crate.
pub type Result<T> = std::result::Result<T, Error>;
