use std::fmt;

use super::Format;

/// Result type for format-neutral iWork reads.
pub type Result<T> = std::result::Result<T, Error>;

/// Stable category of a format-neutral iWork failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ErrorKind {
    /// The bytes are not a recognized Pages, Keynote, or Numbers package.
    Unrecognized,
    /// A caller-selected resource ceiling is invalid.
    InvalidOptions,
    /// An input operation failed.
    Io,
    /// The package or its semantic model is malformed.
    InvalidData,
    /// A finite resource ceiling was exceeded.
    LimitExceeded,
    /// Memory could not be reserved before publishing semantic state.
    Allocation,
    /// An internal invariant was violated.
    Invariant,
}

/// Content-free processing stage associated with an [`Error`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Stage {
    /// Initial byte admission.
    Input,
    /// Application-family detection.
    Detection,
    /// Format-owned semantic decoding.
    Semantic,
    /// Format-neutral snapshot validation.
    Validation,
}

/// Bounded resource associated with an [`Error`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Resource {
    /// Complete packaged input bytes.
    InputBytes,
    /// Number of packaged entries.
    Entries,
    /// Bytes retained by one packaged entry.
    EntryBytes,
    /// Aggregate expanded package bytes.
    ExpandedBytes,
    /// Bytes decoded for one internal package unit.
    DecodedBytes,
    /// Semantic Numbers tables.
    Tables,
    /// Semantic Keynote slides.
    Slides,
    /// Semantic Pages sections.
    Sections,
    /// Materialized Numbers cells.
    Cells,
    /// Aggregate retained UTF-8 text bytes.
    TextBytes,
    /// Other bounded semantic traversal work.
    SemanticWork,
    /// Destination memory.
    Memory,
}

/// A content-free, facade-owned iWork read failure.
///
/// The value deliberately carries no file name, authored text, native object
/// number, lower-layer error, or package implementation type. It is therefore
/// cheap to copy, safe to log, and stable across internal parser migrations.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Error {
    kind: ErrorKind,
    stage: Stage,
    format: Option<Format>,
    resource: Option<Resource>,
    observed: Option<u64>,
    maximum: Option<u64>,
}

impl Error {
    pub(super) const fn new(
        kind: ErrorKind,
        stage: Stage,
        format: Option<Format>,
        resource: Option<Resource>,
        observed: Option<u64>,
        maximum: Option<u64>,
    ) -> Self {
        Self {
            kind,
            stage,
            format,
            resource,
            observed,
            maximum,
        }
    }

    pub(super) const fn unrecognized() -> Self {
        Self::new(
            ErrorKind::Unrecognized,
            Stage::Detection,
            None,
            None,
            None,
            None,
        )
    }

    pub(super) const fn invalid_options(resource: Resource, observed: u64, maximum: u64) -> Self {
        Self::new(
            ErrorKind::InvalidOptions,
            Stage::Validation,
            None,
            Some(resource),
            Some(observed),
            Some(maximum),
        )
    }

    pub(super) const fn invalid_data(format: Option<Format>, stage: Stage) -> Self {
        Self::new(ErrorKind::InvalidData, stage, format, None, None, None)
    }

    pub(super) const fn limit(
        format: Option<Format>,
        stage: Stage,
        resource: Resource,
        observed: u64,
        maximum: u64,
    ) -> Self {
        Self::new(
            ErrorKind::LimitExceeded,
            stage,
            format,
            Some(resource),
            Some(observed),
            Some(maximum),
        )
    }

    pub(super) const fn allocation(format: Option<Format>, stage: Stage, amount: u64) -> Self {
        Self::new(
            ErrorKind::Allocation,
            stage,
            format,
            Some(Resource::Memory),
            Some(amount),
            None,
        )
    }

    pub(super) const fn invariant(format: Option<Format>, stage: Stage) -> Self {
        Self::new(ErrorKind::Invariant, stage, format, None, None, None)
    }

    /// Return the stable failure category.
    #[must_use]
    pub const fn kind(self) -> ErrorKind {
        self.kind
    }

    /// Return the processing stage at which the failure occurred.
    #[must_use]
    pub const fn stage(self) -> Stage {
        self.stage
    }

    /// Return the selected application family when it was known.
    #[must_use]
    pub const fn format(self) -> Option<Format> {
        self.format
    }

    /// Return the bounded resource involved in the failure, when applicable.
    #[must_use]
    pub const fn resource(self) -> Option<Resource> {
        self.resource
    }

    /// Return the observed or requested amount, when applicable.
    #[must_use]
    pub const fn observed(self) -> Option<u64> {
        self.observed
    }

    /// Return the configured maximum, when applicable.
    #[must_use]
    pub const fn maximum(self) -> Option<u64> {
        self.maximum
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.kind {
            ErrorKind::Unrecognized => {
                formatter.write_str("input is not a recognized Apple iWork package")
            },
            ErrorKind::InvalidOptions => write_resource_error(
                formatter,
                "invalid iWork resource ceiling",
                self.resource,
                self.observed,
                self.maximum,
            ),
            ErrorKind::Io => formatter.write_str("iWork input operation failed"),
            ErrorKind::InvalidData => formatter.write_str("invalid Apple iWork data"),
            ErrorKind::LimitExceeded => write_resource_error(
                formatter,
                "iWork resource ceiling exceeded",
                self.resource,
                self.observed,
                self.maximum,
            ),
            ErrorKind::Allocation => formatter.write_str("iWork semantic allocation failed"),
            ErrorKind::Invariant => formatter.write_str("iWork semantic invariant failed"),
        }
    }
}

impl std::error::Error for Error {}

fn write_resource_error(
    formatter: &mut fmt::Formatter<'_>,
    prefix: &str,
    resource: Option<Resource>,
    observed: Option<u64>,
    maximum: Option<u64>,
) -> fmt::Result {
    write!(formatter, "{prefix}")?;
    if let Some(resource) = resource {
        write!(formatter, " for {resource:?}")?;
    }
    if let Some(observed) = observed {
        write!(formatter, ": observed {observed}")?;
    }
    if let Some(maximum) = maximum {
        write!(formatter, ", maximum {maximum}")?;
    }
    Ok(())
}
