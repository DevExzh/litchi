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

    #[error("invalid IWA protobuf header: {0}")]
    Protobuf(#[from] prost::DecodeError),

    #[error("IWA archive reader error: {0}")]
    Io(#[from] std::io::Error),

    #[error("could not encode IWA protobuf header: {0}")]
    ProtobufEncode(#[from] prost::EncodeError),

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
}

/// Result alias for this crate.
pub type Result<T> = std::result::Result<T, Error>;
