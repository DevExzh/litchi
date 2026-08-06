//! Semantic views over source-preserved VBA signature blobs.

use std::fmt;
use std::ops::Range;
use std::sync::Arc;

use super::codec;

/// The two `[MS-OSHARED]` containers for serialized VBA signature metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// A `DigSigBlob`, as embedded by the `VtDigSig` property value.
    Property,
    /// A `WordSigBlob`, whose outer size is expressed as UTF-16 code units.
    Word,
}

/// Resource limits applied before retaining an input allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum complete serialized blob size.
    pub max_blob_bytes: usize,
    /// Maximum opaque signature payload size.
    pub max_signature_bytes: usize,
    /// Maximum opaque serialized certificate-store size.
    pub max_certificate_store_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_blob_bytes: 16 * 1024 * 1024,
            max_signature_bytes: 8 * 1024 * 1024,
            max_certificate_store_bytes: 8 * 1024 * 1024,
        }
    }
}

/// A malformed or resource-exhausting serialized signature blob.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    /// The source ended before a declared field boundary.
    Truncated(&'static str),
    /// A wire invariant required by `[MS-OSHARED]` was violated.
    Invalid(String),
    /// A configured resource budget was exceeded.
    Limit(&'static str),
}

impl Error {
    pub(crate) fn invalid(message: impl Into<String>) -> Self {
        Self::Invalid(message.into())
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Truncated(field) => write!(formatter, "truncated VBA signature {field}"),
            Self::Invalid(message) => write!(formatter, "invalid VBA signature blob: {message}"),
            Self::Limit(resource) => write!(formatter, "VBA signature {resource} limit exceeded"),
        }
    }
}

impl std::error::Error for Error {}

/// A validated, exact-source VBA signature serialization.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Blob {
    source: Arc<[u8]>,
    kind: Kind,
    layout: codec::Layout,
}

impl Blob {
    /// Parses a complete blob using conservative default resource limits.
    pub fn parse(source: &[u8], kind: Kind) -> Result<Self, Error> {
        Self::parse_with(source, kind, Limits::default())
    }

    /// Parses a complete blob using caller-provided resource limits.
    pub fn parse_with(source: &[u8], kind: Kind, limits: Limits) -> Result<Self, Error> {
        let layout = codec::parse(source, kind, limits)?;
        Ok(Self {
            source: Arc::from(source),
            kind,
            layout,
        })
    }

    /// Parses an already-shared source allocation without copying it.
    pub fn parse_shared(source: Arc<[u8]>, kind: Kind, limits: Limits) -> Result<Self, Error> {
        let layout = codec::parse(&source, kind, limits)?;
        Ok(Self {
            source,
            kind,
            layout,
        })
    }

    /// Parses a complete `DigSigBlob` using default limits.
    pub fn parse_property(source: &[u8]) -> Result<Self, Error> {
        Self::parse(source, Kind::Property)
    }

    /// Parses a complete `WordSigBlob` using default limits.
    pub fn parse_word(source: &[u8]) -> Result<Self, Error> {
        Self::parse(source, Kind::Word)
    }

    /// Returns the source container kind.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Borrows the typed nested signature information.
    #[must_use]
    pub const fn info(&self) -> Info<'_> {
        Info { blob: self }
    }

    /// Returns the exact source bytes, including opaque gaps and padding.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Clones ownership of the exact source allocation without copying bytes.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.source)
    }

    /// Moves the exact source allocation out of this value.
    #[must_use]
    pub fn into_bytes(self) -> Arc<[u8]> {
        self.source
    }
}

/// A borrowed semantic view of `DigSigInfoSerialized`.
#[derive(Debug, Clone, Copy)]
pub struct Info<'a> {
    blob: &'a Blob,
}

impl<'a> Info<'a> {
    /// Returns the opaque VBA signature payload.
    #[must_use]
    pub fn signature(self) -> &'a [u8] {
        self.slice(&self.blob.layout.signature)
    }

    /// Returns the opaque serialized certificate-store payload.
    #[must_use]
    pub fn certificate_store(self) -> &'a [u8] {
        self.slice(&self.blob.layout.certificate_store)
    }

    /// Returns the reserved project-name field in its exact wire form.
    ///
    /// A conforming value is one null UTF-16 code unit.
    #[must_use]
    pub fn reserved_project_name(self) -> &'a [u8] {
        self.slice(&self.blob.layout.project_name)
    }

    /// Returns the reserved timestamp-URL field in its exact wire form.
    ///
    /// A conforming value is one null UTF-16 code unit.
    #[must_use]
    pub fn reserved_timestamp_url(self) -> &'a [u8] {
        self.slice(&self.blob.layout.timestamp_url)
    }

    /// Returns the reserved timestamp marker exactly as read.
    ///
    /// `[MS-OSHARED]` directs consumers to ignore this value on read, so it is
    /// surfaced and preserved rather than normalized.
    #[must_use]
    pub const fn reserved_timestamp_marker(self) -> u32 {
        self.blob.layout.timestamp_marker
    }

    /// Returns the undefined alignment bytes following the signature info.
    #[must_use]
    pub fn padding(self) -> &'a [u8] {
        self.slice(&self.blob.layout.padding)
    }

    /// Returns the complete nested `DigSigInfoSerialized` byte range.
    #[must_use]
    pub fn bytes(self) -> &'a [u8] {
        self.slice(&self.blob.layout.info)
    }

    fn slice(self, range: &Range<usize>) -> &'a [u8] {
        &self.blob.source[range.clone()]
    }
}
