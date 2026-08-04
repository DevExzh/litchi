//! Bounded physical-layer primitives for Apple iWork IWA components.
//!
//! This crate owns the checksum-free Snappy framing used by `.iwa` members and
//! the copyable resource budgets that higher-level archive parsers can carry
//! across API boundaries. It deliberately does not open packages or interpret
//! application-specific protobuf messages.

#![forbid(unsafe_code)]

pub mod archive;
mod error;
mod limits;
mod snappy;

pub use archive::{Archive, ArchiveInfo, ArchiveObject, MessageInfo, RawMessage};
pub use error::{Error, LimitKind, Result};
pub use limits::Limits;
pub use snappy::{SnappyLimits, SnappyStream};

/// The archive/resource limit profile used by IWA parsers.
pub type ArchiveLimits = Limits;

/// The archive/resource limit profile under its shorter generic name.
pub type ResourceLimits = Limits;
