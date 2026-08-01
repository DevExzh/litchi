//! Inert RTF protection-exception range markers.
//!
//! The RTF 1.9.1 specification (Word 2003 document protection) defines the
//! `\*\protstart` and `\*\protend` destinations, which delimit ranges of body
//! text excluded from document protection. Each destination carries an opaque
//! hexadecimal identifier; matching identifiers pair a start with an end, and
//! ranges may overlap arbitrarily (they are not required to nest). The
//! companion `\*\protusertbl` user table is modeled in
//! [`crate::ProtectionUserTable`].
//!
//! The markers are parsed and stored as passive metadata only: identifiers
//! are kept verbatim, and no editing restriction is ever evaluated or
//! enforced.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_PROTECTION_RANGES: usize = 65_536;
pub(crate) const MAX_PROTECTION_RANGE_ID_BYTES: usize = 64;

/// One inert protection-exception range spanning a range of body text.
///
/// The range starts at `position` (a UTF-8 byte offset into the document
/// body text) and covers `content`, mirroring how [`crate::Bookmark`] ranges
/// are recorded. Ranges parsed from a document are ordered by the source
/// order of their `\*\protstart` markers.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionRange<'a> {
    /// Opaque hexadecimal identifier pairing the `\*\protstart` and
    /// `\*\protend` markers, stored verbatim.
    pub id: Cow<'a, str>,
    /// UTF-8 byte offset in the document body text where the range starts.
    pub position: usize,
    /// Body text covered by the range.
    pub content: Cow<'a, str>,
}

impl<'a> ProtectionRange<'a> {
    /// Create a validated protection-exception range.
    pub fn new(id: Cow<'a, str>, position: usize, content: Cow<'a, str>) -> RtfResult<Self> {
        let range = Self {
            id,
            position,
            content,
        };
        range.validate()?;
        Ok(range)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.id.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF protection-range identifier cannot be empty".to_string(),
            ));
        }
        if self.id.len() > MAX_PROTECTION_RANGE_ID_BYTES
            || !self.id.len().is_multiple_of(2)
            || !self.id.bytes().all(|byte| byte.is_ascii_hexdigit())
        {
            return Err(RtfError::MalformedDocument(
                "RTF protection-range identifier must be even-length hexadecimal within the safety limit"
                    .to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> ProtectionRange<'static> {
        ProtectionRange {
            id: Cow::Owned(self.id.into_owned()),
            position: self.position,
            content: Cow::Owned(self.content.into_owned()),
        }
    }
}
