//! Inert generated list-marker destinations (`listtext` and `pntext`).

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_GENERATED_LIST_MARKERS: usize = 65_536;
pub(crate) const MAX_GENERATED_LIST_MARKER_BYTES: usize = 4_096;
pub(crate) const MAX_GENERATED_LIST_MARKER_TOTAL_BYTES: usize = 4 * 1_048_576;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GeneratedListMarkerKind {
    Modern,
    Legacy,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeneratedListMarker<'a> {
    pub kind: GeneratedListMarkerKind,
    pub text: Cow<'a, str>,
    /// UTF-8 byte position in visible document text.
    pub position: usize,
}

impl GeneratedListMarker<'_> {
    pub fn validate(&self) -> RtfResult<()> {
        if self.text.is_empty() || self.text.len() > MAX_GENERATED_LIST_MARKER_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF generated list marker is empty or exceeds the safety limit".to_string(),
            ));
        }
        if self
            .text
            .chars()
            .any(|character| character == '\0' || (character.is_control() && character != '\t'))
        {
            return Err(RtfError::MalformedDocument(
                "RTF generated list marker contains a forbidden control character".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> GeneratedListMarker<'static> {
        GeneratedListMarker {
            kind: self.kind,
            text: Cow::Owned(self.text.into_owned()),
            position: self.position,
        }
    }
}
