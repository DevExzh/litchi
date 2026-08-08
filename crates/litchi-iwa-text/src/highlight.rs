//! Archive-free semantic values for native iWork text highlights.
//!
//! Native annotation lookup, protobuf decoding, and package mutation remain
//! in the owning IWA adapter. This leaf owns only the compact checked identity
//! and UTF-16 range exchanged by the format crates.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Opaque IDs keep their raw adapter module adjacent to private representations."
)]
#![allow(
    clippy::module_name_repetitions,
    reason = "TextHighlight names identify the semantic value in this focused module."
)]

use std::num::NonZeroU64;

use crate::position::TextRange;

/// Validation failures produced while constructing text-highlight values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The highlight identity is zero.
    ZeroId,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ZeroId => formatter.write_str("iWork text-highlight identifier must be non-zero"),
        }
    }
}

impl std::error::Error for Error {}

/// Result type for text-highlight semantic values.
pub type Result<T> = std::result::Result<T, Error>;

/// A compact, non-zero identifier for a native text highlight.
//
// Keeping the invalid zero state out of the value preserves the native
// eight-byte representation while making `Option<TextHighlightId>` compact.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextHighlightId(NonZeroU64);

/// Explicit native-boundary conversions for the opaque highlight handle.
pub mod raw {
    use super::{Result, TextHighlightId};

    /// Validate a native highlight object identifier at the IWA boundary.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::ZeroId`] when `identifier` is zero.
    pub const fn from_object_id(identifier: u64) -> Result<TextHighlightId> {
        TextHighlightId::from_raw(identifier)
    }

    /// Recover a native highlight object identifier inside an adapter.
    #[must_use]
    pub const fn object_id(identifier: TextHighlightId) -> u64 {
        identifier.into_raw()
    }
}

impl TextHighlightId {
    const fn from_raw(identifier: u64) -> Result<Self> {
        match NonZeroU64::new(identifier) {
            Some(non_zero) => Ok(Self(non_zero)),
            None => Err(Error::ZeroId),
        }
    }

    const fn into_raw(self) -> u64 {
        self.0.get()
    }
}

/// One native plain highlight attached to a nonempty UTF-16 text range.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TextHighlight {
    /// Stable semantic identity used by the owning adapter for updates.
    pub id: TextHighlightId,
    /// Half-open UTF-16 highlighted text range.
    pub range: TextRange,
}

impl TextHighlight {
    /// Construct a highlight from its validated semantic components.
    #[must_use]
    pub const fn new(id: TextHighlightId, range: TextRange) -> Self {
        Self { id, range }
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;

    #[test]
    fn identifiers_are_nonzero_and_compact() {
        assert_eq!(raw::from_object_id(0), Err(Error::ZeroId));
        let identifier = raw::from_object_id(42).unwrap();
        assert_eq!(raw::object_id(identifier), 42);
        assert_eq!(size_of::<TextHighlightId>(), size_of::<u64>());
        assert_eq!(
            size_of::<Option<TextHighlightId>>(),
            size_of::<TextHighlightId>()
        );
    }
}
