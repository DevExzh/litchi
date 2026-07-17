//! Strict public types for native iWork plain-text highlights.

use crate::{Error, Result};

use super::position::TextRange;

/// Identifier of a native text-highlight object.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextHighlightId(u64);

impl TextHighlightId {
    /// Construct an identifier obtained from a previously read highlight.
    pub fn from_object_id(identifier: u64) -> Result<Self> {
        if identifier == 0 {
            return Err(Error::ParseError(
                "iWork text-highlight object identifier cannot be zero".to_owned(),
            ));
        }
        Ok(Self(identifier))
    }

    /// Return the underlying package object identifier.
    pub const fn object_id(self) -> u64 {
        self.0
    }

    pub(crate) const fn from_native(identifier: u64) -> Self {
        Self(identifier)
    }
}

/// One native plain highlight attached to a nonempty UTF-16 text range.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TextHighlight {
    /// Stable highlight-object identifier used for update and deletion.
    pub id: TextHighlightId,
    /// Half-open highlighted text range.
    pub range: TextRange,
}

impl TextHighlight {
    pub(crate) const fn new(id: TextHighlightId, range: TextRange) -> Self {
        Self { id, range }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn zero_identifier_is_rejected() {
        assert!(TextHighlightId::from_object_id(0).is_err());
        assert_eq!(TextHighlightId::from_object_id(42).unwrap().object_id(), 42);
    }
}
