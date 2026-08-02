//! Inert RTF editable-region boundary markers.
//!
//! Microsoft Word writes the parameterless `\ebcstart` and `\ebcend` control
//! words inline in the body text of protected documents to delimit regions
//! that remain editable (the RTF counterpart of OOXML
//! `w:permStart`/`w:permEnd`). The marks carry no identifier, so they pair
//! positionally: each `\ebcend` closes the innermost open `\ebcstart`, and
//! regions must be properly nested.
//!
//! The markers are parsed and stored as passive metadata only; no editing
//! restriction is ever evaluated or enforced.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_EDITABLE_REGIONS: usize = 65_536;

/// One inert editable region spanning a range of body text.
///
/// The region starts at `position` (a UTF-8 byte offset into the document
/// body text) and covers `content`, mirroring how [`crate::Bookmark`] ranges
/// are recorded. Regions parsed from a document are ordered by the source
/// order of their `\ebcstart` marks and are properly nested.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EditableRegion<'a> {
    /// UTF-8 byte offset in the document body text where the region starts.
    pub position: usize,
    /// Body text covered by the region.
    pub content: Cow<'a, str>,
}

impl<'a> EditableRegion<'a> {
    /// Create a validated editable region.
    pub fn new(position: usize, content: Cow<'a, str>) -> RtfResult<Self> {
        let region = Self { position, content };
        region.validate()?;
        Ok(region)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.content.contains('\0') {
            return Err(RtfError::MalformedDocument(
                "RTF editable-region content contains a forbidden control character".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> EditableRegion<'static> {
        EditableRegion {
            position: self.position,
            content: Cow::Owned(self.content.into_owned()),
        }
    }
}
