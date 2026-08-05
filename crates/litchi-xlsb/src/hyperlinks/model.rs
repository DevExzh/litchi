//! Semantic `BrtHLink` values for XLSB worksheets.

use thiserror::Error;

/// The fixed-width `rfx` prefix of a `BrtHLink` payload.
pub const PREFIX_LEN: usize = 16;

/// Result type for hyperlink parsing and serialization.
pub type Result<T> = std::result::Result<T, Error>;

/// A typed `BrtHLink` failure.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The payload does not contain the fixed-width `rfx` prefix.
    #[error("invalid BrtHLink payload length: expected at least {expected} bytes, found {found}")]
    InvalidLength {
        /// Minimum payload length.
        expected: usize,
        /// Actual payload length.
        found: usize,
    },
    /// A scalar or string failed validated BIFF12 decoding or encoding.
    #[error(transparent)]
    Wire(#[from] crate::raw::Error),
}

/// Hyperlink information for a cell or range of cells.
///
/// The relationship target is intentionally kept as an optional writer-side
/// value; resolving package relationships is a host concern.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    /// First row (zero-based).
    pub row_first: u32,
    /// Last row (zero-based, inclusive).
    pub row_last: u32,
    /// First column (zero-based).
    pub col_first: u32,
    /// Last column (zero-based, inclusive).
    pub col_last: u32,
    /// Relationship ID in the worksheet relationship part.
    pub r_id: String,
    /// Location within the destination document or workbook.
    pub location: Option<String>,
    /// Tooltip text.
    pub tooltip: Option<String>,
    /// Display text.
    pub display: Option<String>,
    /// External hyperlink target URL (writer-side only).
    pub target: Option<String>,
}

impl Hyperlink {
    /// Create a hyperlink with an explicit relationship ID.
    #[must_use]
    pub fn new(row_first: u32, row_last: u32, col_first: u32, col_last: u32, r_id: String) -> Self {
        Self {
            row_first,
            row_last,
            col_first,
            col_last,
            r_id,
            location: None,
            tooltip: None,
            display: None,
            target: None,
        }
    }

    /// Create an internal hyperlink pointing to a workbook location.
    #[must_use]
    pub fn new_internal(
        row_first: u32,
        row_last: u32,
        col_first: u32,
        col_last: u32,
        location: String,
    ) -> Self {
        Self {
            row_first,
            row_last,
            col_first,
            col_last,
            r_id: String::new(),
            location: Some(location),
            tooltip: None,
            display: None,
            target: None,
        }
    }

    /// Create an external hyperlink pointing to a URL.
    #[must_use]
    pub fn new_external(
        row_first: u32,
        row_last: u32,
        col_first: u32,
        col_last: u32,
        target: String,
    ) -> Self {
        Self {
            row_first,
            row_last,
            col_first,
            col_last,
            r_id: String::new(),
            location: None,
            tooltip: None,
            display: None,
            target: Some(target),
        }
    }

    /// Set a fragment or workbook location.
    #[must_use]
    pub fn with_location(mut self, location: String) -> Self {
        self.location = Some(location);
        self
    }

    /// Set tooltip text.
    #[must_use]
    pub fn with_tooltip(mut self, tooltip: String) -> Self {
        self.tooltip = Some(tooltip);
        self
    }

    /// Set display text.
    #[must_use]
    pub fn with_display(mut self, display: String) -> Self {
        self.display = Some(display);
        self
    }

    /// Parse a `BrtHLink` payload.
    #[inline]
    pub fn parse(data: &[u8]) -> Result<Self> {
        super::codec::parse(data)
    }

    /// Serialize a `BrtHLink` payload with checked resource limits.
    #[inline]
    pub fn try_serialize(&self) -> Result<Vec<u8>> {
        super::codec::serialize(self)
    }

    /// Serialize to a `BrtHLink` payload using the historical infallible
    /// writer. Untrusted callers should prefer [`Self::try_serialize`].
    #[must_use]
    pub fn serialize(&self) -> Vec<u8> {
        super::codec::serialize_legacy(self)
    }
}
